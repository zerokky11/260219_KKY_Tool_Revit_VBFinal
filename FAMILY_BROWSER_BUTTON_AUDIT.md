# KKY Family Browser Button Audit

Last updated: 2026-07-22 11:23 KST

## Working Rule

- Always read this file before continuing the audit.
- Treat this file as the source of truth for what has been checked, what was fixed, and what remains.
- After every meaningful scan, finding, code change, build, install, or verification, update this file.
- Do not mark an item as done unless the code path was traced to its handler and the expected state changes were checked.
- If a finding needs Revit runtime confirmation, mark it as `Needs Revit Check` instead of guessing.
- Before editing source files, create a source backup and record the backup path here.

## Scope

- Workspace: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628`
- Main dashboard files:
  - `KKY_FamilyBrowser_RevitHost_2019-2023\FamilyBrowserDashboardHtmlForm.cs`
  - `KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserDashboardHtmlForm.cs`
  - `KKY_FamilyBrowser_RevitHost_2027\FamilyBrowserDashboardHtmlForm.cs`
- Target Revit versions: 2019, 2021, 2023, 2025, 2027
- Add-in install targets:
  - `C:\ProgramData\Autodesk\Revit\Addins\2019`
  - `C:\ProgramData\Autodesk\Revit\Addins\2021`
  - `C:\ProgramData\Autodesk\Revit\Addins\2023`
  - `C:\ProgramData\Autodesk\Revit\Addins\2025`
  - `C:\ProgramData\Autodesk\Revit\Addins\2027`

## Audit Method

1. Extract all visible and generated actions:
   - `kkyfb:*`
   - JavaScript `window.location`
   - dashboard tab links
   - standard RVT manager window actions
   - detail window actions
   - selection dialog actions
2. Map every action to:
   - generating UI code
   - router branch
   - handler method
   - permission check
   - required state
   - files/policy records read
   - files/policy records written
   - cache invalidation and dashboard refresh path
3. Compare the traced behavior against intended behavior.
4. Record each item in the Button Audit Table.
5. Fix clear bugs after backup, then build/install/verify all target versions.

## Status Legend

- `Not Started`: not traced yet
- `Tracing`: currently being followed through code
- `OK`: traced and behavior matches intent
- `Fixed`: bug found and patched
- `Needs Design`: code works but intended behavior needs user decision
- `Needs Revit Check`: code path looks right, but actual Revit document/API behavior must be confirmed inside Revit
- `Blocked`: cannot proceed without missing file/tool/state

## Automation Status

- Requested interval: 1 minute
- Automation style: thread heartbeat
- Automation id: `kky-family-browser-button-audit-heartbeat`
- Automation status: DELETED on 2026-06-29 16:06 KST at user request.
- No automatic heartbeat should continue after this point.
- Resume manually by reading this md first, then following `Tomorrow Resume Notes` and `Next Work Queue`.

## Tomorrow Resume Notes

- Latest File Guard trade/automatic-check implementation: every guarded RVT now stores one assigned standard trade, the HTML configuration window supports per-row and bulk trade selection, and Korean/English XLSX export/import includes `공종` / `Discipline`. On first project open, the add-in schedules a non-modal Current Model Check against that assigned trade, reuses a valid saved comparison, serializes concurrent sessions with a project lock, and creates no automatic workbook. Nested-only placement candidates and fingerprint refreshes are now scoped to the same assigned trade instead of aggregating every registered standard. Backup `_backups\file-trade-auto-model-check-20260721-174639`; ProgramData backup `_backups\programdata-before-file-trade-auto-model-check-20260721-182931`. Five-target build/Stage/installed verification and focused Korean/English IE checks passed; real Revit first-open scan/cache/multi-PC behavior remains Next Work Queue item 43.
- Latest project catalog tracking: browser startup/activation, explicit Refresh, and successful Save/Save As/Synchronize now capture a per-project name-only catalog containing loadable Family names, FamilySymbol type names, and supported System Type names. Canonical project identity uses the central-model path when available. Accepted baseline and last-observed state are separate, so an unresolved warning cannot disappear merely because the browser observed the same changed model again.
- Latest project catalog UI/history: Home and the header show persistent baseline-missing or added/removed warnings. The Home table classifies deltas matched by committed Browser `OperationLogs` as known Browser changes and unmatched additions/removals as `외부/미추적 변경`; no external author or exact mutation time is inferred. A successful Current Model Check accepts the current catalog, and administrators can explicitly run `변경 확인 및 기준 갱신` after reviewing the name-only warning.
- Latest project catalog verification/deployment: source backup `_backups\project-catalog-tracking-20260718-154307`; ProgramData backup `_backups\programdata-before-project-catalog-20260718-162416`. Targeted Korean/English IE report `artifacts\family-browser-ui-audit\20260718-162051` passed all project-catalog scenarios for installed Revit runtimes; 2021 remains `SKIP runtime-not-installed`. Integrated no-harness report `artifacts\family-browser-ui-audit\20260718-162230-quality-gate` passed static/contract, nested-family propagation, authoritative System Type apply, five-target Stage verification, and 2,000-row performance/cache. ProgramData was updated for 2019/2021/2023/2025/2027 and all Stage/installed DLL hashes match.
- Latest Standard RVT revision/audit fix: accepted scans now record an atomic per-source revision manifest containing path/file identity, timestamp, length, and `SHA256-SAMPLED-V1`. Browser startup, top Refresh, and the 60-second background probe compare the registered source without opening the RVT. Changed, missing, inaccessible, or baseline-missing sources produce a persistent per-trade Home board/header warning, clear stale Family/System rows, and block model check/load/apply until an administrator accepts a new scan. Mapped drive/UNC/final-path/hardlink aliases are canonicalized through Windows file identity where available.
- Latest Standard RVT history fix: pending edits are committed to immutable managed history only after successful Save, Save As, or Synchronize. Entries include Revit/Windows user, machine, source identity, category/family/type, Added/Modified/Deleted, before/after fingerprint summaries, commit kind/time, and are shown in Standards. An external replacement is detected and blocked but cannot truthfully reconstruct its author or exact changed items.
- Latest verification/deployment: backup `_backups\standard-rvt-revision-history-20260718-082156`. Full report `artifacts\family-browser-ui-audit\20260718-134334-standard-revision-quality-gate` passed static/contract, nested-family propagation, authoritative System Type apply, five-target Stage verification, 2,000-row performance/cache, and all Korean/English IE DOM/click scenarios with zero failures. Revit 2021 runtime remains `SKIP runtime-not-installed`; its build/Stage/ProgramData/package passed. ProgramData Stage hashes match for all five targets.
- Latest installer/mail package: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260718-140036-standard-rvt-tracking_Setup.exe`, 3.45 MB, SHA256 `EC23A9BA5A36E46AE64AA31779F510364F70ACFF3F8CF1F946AC6397F37A5A99`; mail package `artifacts\family-browser\mail-packages\20260718_01.zip`, 16.00 MB, SHA256 `C6BC781F1BD61DA2DFE4F823A0EBFA5BB25DD0B12CA2F188E723AB3F647F519E`. ZIP `Setup.exe` hash matches the standalone installer.
- Latest authoritative System Type apply fix: code tracing found that compound-layer signatures were captured/compared but ordinary create/overwrite did not guarantee a `SetCompoundStructure(...)` mutation after canonical copy. It also found `RailingType` and `StairsType` configured as `ReviewOnly`, so they were skipped by apply. All three host source sets now call the standard-definition apply after every create/copy path; map compound layer materials, deck profiles, and wall sweeps into the target document; call `SetCompoundStructure`; and fail/rollback when `CompoundStructure.IsEqual` post-check differs. Same-name materials are reused and synchronized, missing materials are created, and positive-ID guards preserve Revit `By Category`/empty references without creating duplicate names. If a same-name layer/wall-sweep material still differs after synchronization, the whole item rolls back instead of silently keeping a partially matching material.
- Latest Railing/Stair apply fix: `RailingType` and `StairsType` now use `PreflightThenConfirm`. Apply refreshes referenced loadable families, resolves referenced system types by exact class/category/name, synchronizes writable parameters/API properties, preserves readonly nested structure through canonical root copy, and compares source/target detailed signatures after mutation. This path is intentionally independent of `CompareDetailedSystemTypeComponents`; the option still controls comparison visibility/cost only. Curtain Panel remains review-only.
- Latest verification/deployment: backup `_backups\system-type-authoritative-apply-20260716-131553`. New gate `KKY_FamilyBrowser_Compile\Test-SystemTypeAuthoritativeApply.ps1` is included in the integrated quality gate. Static/contract, dedicated regression, zero-error Release compilation, five-target Stage verification, 2,000-row performance/cache, and IE WebBrowser harness passed. Harness completed 38 scenarios each for 2019/2023/2025/2027 with zero failures; 2021 remains `SKIP runtime-not-installed`. ProgramData was updated for 2019/2021/2023/2025/2027 and Stage/installed SHA256 values match. Report: `artifacts\family-browser-ui-audit\20260716-system-type-authoritative-apply-final\quality-gate-summary.md`.
- Latest installer/mail package: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260716-140154_Setup.exe`, SHA256 `EF88DBF6301895B4552FC998A19685D085A4867BF378AF9A96657D8C135894DC`; mail package `artifacts\family-browser\mail-packages\20260716_04.zip`, `16,777,451` bytes, SHA256 `28261CE9B40BCD37D06E9533DBECE6E9E7160E9BA52912CCD2227CA9EEE8CC07`. Both contain Revit 2019/2021/2023/2025/2027 and include the authoritative System Type apply revision.
- Needs Revit Check: use a disposable blank RVT to apply and inspect multi-layer Wall/Floor/Roof/Ceiling types, including materials/core/variable/structural/wrapping/deck/wall-sweep data. Test Railing/Stair create and overwrite with detailed comparison ON and OFF, verify all referenced component types/native dialog values, and confirm no suffixed material/family/system-type names are produced.
- Latest nested-only placement policy: File Guard now has an explicit per-RVT `하위 전용 패밀리 단독 모델링 금지` option. A family is eligible only when the selected standard's latest precise scan records it as a nested child and records zero project-level standalone instances. Parent placement and its generated nested instances remain allowed; a newly added `FamilyInstance` with no `SuperComponent` is blocked before commit while Admin Mode is OFF. Legacy/incomplete scans fail open and require a new precise scan. Sidecar/cache: `<snapshot>.nested-only-placement-v1.json` and `%LOCALAPPDATA%\KKY\FamilyBrowser\Cache\Guard\NestedOnly`; full/selected scan attachment now invalidates this catalog cache immediately. Backup: `_backups\nested-only-placement-guard-20260715-172727`. Dedicated policy: `FAMILY_BROWSER_NESTED_ONLY_PLACEMENT_POLICY.md`. Final report: `artifacts\family-browser-ui-audit\20260715-nested-only-placement-final-v3\quality-gate-summary.md`; static/contract, five-target Release build/Stage, 2,000-row performance/cache, and `136 OK + 1 expected 2021 runtime SKIP` passed with zero failures. ProgramData and installer artifacts were not changed. Real Revit placement rollback remains queue item 28.
- Latest compound-layer detail fix: System Type detail now renders compound layers as a Revit Edit Assembly-style table with exterior/interior direction, index, function, material, thickness, core boundaries, structural-material badge, and variable-layer badge. Routing criteria and layer thickness share synchronized `mm` / `in` selectors; `mm` is the first-use default and the last choice is restored from `%LOCALAPPDATA%\KKY\FamilyBrowser\Settings\measurement-unit.txt` on the next browser run. Existing legacy snapshots still render through a compatibility parser; a new precise scan is required to populate core/structural/variable metadata. Backups: `_backups\system-layer-unit-ui-20260715-160848` and `_backups\system-detail-hidden-block-qa-20260715-165016`. Final visual QA removed an audit-only CSS rule that incorrectly exposed hidden system-detail blocks and added a production `fb-system-detail` guard, so the detached System detail cannot show the Family composition or legacy bottom preview card. Final report: `artifacts\family-browser-ui-audit\20260715-system-layer-unit-final-v2\quality-gate-summary.md`; static/contract, five-target build/Stage, 2,000-row performance/cache, `136 OK + 1 expected 2021 runtime SKIP`, and generated system-detail HTML/PNG checks passed with zero failures. ProgramData and installer artifacts were not changed.
- Latest Family/System trade-switch fix: clicking a prepared trade chip below the search field now changes the actual Family/System dataset instead of leaving rows from the Standards-selected trade visible. The switch reuses the prepared target slot, restores that target's saved project comparison on demand before the row-cache key is calculated, and shows immediate active-chip/old-row-hide feedback while the host refresh completes. Standard RVT/list mutations invalidate prepared slots so the acceleration cannot serve stale data. Backup: `_backups\family-browser-trade-switch-20260715-153052`. Final report: `artifacts\family-browser-ui-audit\20260715-trade-switch-final\quality-gate-summary.md`; static/contract, five-target build/Stage, 2,000-row performance/cache, and `136 OK + 1 expected 2021 runtime SKIP` passed with zero failures. ProgramData and installer artifacts were not changed.
- Latest load/apply commit-lifecycle fix: a successful Family load or System Type apply now immediately replaces `Load Available` with a non-selectable temporary state (`로드됨 · 저장/동기화 대기` / `적용됨 · 저장/동기화 대기`). It is not treated as a persisted completion until Revit reports a successful Save, Save As, or Synchronize with Central. Failed/cancelled commits keep the temporary state, while a completed close without a successful commit discards it so reopening the unsaved model naturally returns to the persisted project state.
- Latest commit-lifecycle verification: `artifacts\family-browser-ui-audit\20260713-pending-save-sync-final` passed static/contract, five-target build/stage, stage verification, 2,000-row performance/cache, and the full IE HTML/click/language/layout harness. Harness result: `136 OK + Revit 2021 runtime-not-installed SKIP`, zero failures. ProgramData was not replaced (`Install: False`). Backup: `_backups\pending-save-sync-lifecycle-20260713-153915`.
- Latest list usability fix: Family/System list fixed headers now expose drag handles on every column boundary. Dragging updates matching header/body `colgroup` widths together, keeps a sensible per-column minimum, and persists widths by tab/column-count through guarded local storage with an in-memory fallback.
- Latest paging fix: the old 150-row arrow/range widget was removed from the search line. When filtered results exceed 150 rows, the bottom screen-status bar now shows localized Previous/Next actions, current/total page summary, clickable numeric pages, ellipsis for long page sets, and the visible row range. Existing cross-page checked rows remain intact.
- Latest list automation/deployment: static QA locks the separated status text/pager markup, direct page API, persisted column-width runtime, 50px status area, and resize seam. The IE `WebBrowser` 2,000-row harness directly opens page 2, verifies active numeric state/summary, keeps page-1/page-2 checks, and changes one header/body column to the same pixel width. Full report `artifacts\family-browser-ui-audit\20260713-100935-quality-gate` passed all seven stages with zero failures; 2021 runtime remains the expected `SKIP runtime-not-installed` while build/install verification passed.
- Latest list installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_resizable-columns-numbered-pager-20260713_Setup.exe`, SHA256 `BB361BDC2FC1301225426A63CFD1BCEC3BF2E13B30AB350FFA74FBA13E1976E8`. Mail package: `artifacts\family-browser\mail-packages\20260713_04.zip`, 15.5 MB, SHA256 `FC79AF599902F49C2AA57FFCC846E8205E785CF962B9763686BA20E0D6B63DF6`.
- Latest list backup: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\resizable-columns-numbered-pager-20260713-095810`.
- Latest full-HTML dialog fix: the previous conversion covered only the message/result body. The visible caption, subtitle, close button, OK/Yes/No buttons, copy button, and operation-result footer were still WinForms `Label` / `Button` controls, which explains the broken title/action text at workstation DPI. `KKY_FamilyBrowser_SharedUi\FamilyBrowserHtmlDialogHost.cs` now leaves one borderless WinForms host containing one full-dock `WebBrowser`; all visible chrome and actions render in the same IE-compatible HTML document.
- Latest dialog scope: shared dashboard messages, scan completion/errors, Yes/No confirmations, family load confirmation/result, and system type apply confirmation/result now use the same full-HTML title/body/footer shell. HTML routes preserve OK=`DialogResult.OK`, affirmative=`Yes`, negative/close/Esc=`No`, Enter default action, details/path copy, report-folder open, title-bar drag, and large-result maximize/restore behavior. A native `MessageBox` remains only as an emergency fallback if constructing the HTML host itself throws.
- Latest dialog automation/deployment: static QA rejects the removed WinForms title/action shells. Korean/English structured-message IE scenarios require `data-dialog-shell=full-html` and semantic title, close, body, footer, copy, and accept elements. `artifacts\family-browser-ui-audit\20260713-full-html-dialog-quality-gate` passed all seven gate stages with 112 `OK`, one Revit 2021 `SKIP runtime-not-installed`, and zero failures. 2019/2021/2023/2025/2027 Stage and ProgramData hashes match.
- Latest dialog installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_full-html-dialog-20260713_Setup.exe`, SHA256 `2DFAB5B62D41E49532FF90680FD0BC23C64D1EED275E07BB9148ED89BF8D47B8`. Mail package: `artifacts\family-browser\mail-packages\20260713_03.zip`, 15.5 MB, SHA256 `28F4F5381F11281239CF5FA690585CB806B2ACFEF2D7BF29AEE217B77F66B85E`.
- Latest dialog backup: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\full-html-dialog-shell-20260713-092523`.
- Latest live-state refresh fix: post-action `RefreshDocumentShellOnly(...)` was reusing `_startupPreloadResult.Policy` and prepared registration/snapshot/list data from the first browser open. Persisted changes were correct on disk, but the immediate render could overwrite them with the startup copy until the user pressed Refresh. Startup preload is now explicitly allowed only for the first shell render; every later fast refresh reloads the live policy and current registration/list metadata while preserving tab/filter/scroll UI state.
- Latest live-state automation: static QA now locks the initial-only preload flag, live-store default, prepared-slot gate, `finally` reset, and diagnostic source. It also requires immediate refresh calls in 18 persisted-state mutation methods covering standard mode/target add/rename/delete, standard list connect/clear, security, permission Excel, file guard, project policy, standard RVT reset, and request-store changes.
- Latest live-state verification/deployment: `artifacts\family-browser-ui-audit\20260713-live-mutation-refresh-quality-gate` passed static/contract, 2019/2021/2023/2025/2027 build/stage, staged verification, 2,000-row performance/cache, 112 IE DOM/click/language/layout/detail scenarios, ProgramData install, and installed verification. Revit 2021 has the expected one runtime skip; failures are zero. New installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_live-state-refresh-20260713_Setup.exe`; mail package: `artifacts\family-browser\mail-packages\20260713_02.zip` (15.5 MB).
- Latest default-window fix: Family Browser now uses the same baseline sizing rule as KKY Tool: desired `1400 x 900`, capped to `93%` of the resolved Revit screen working area, minimum `1100 x 720` without exceeding the actual startup size, and centered on the resolved Revit monitor. The previous wide-screen rule could start at `1680 x 960` or `1780 x 1040` and has been removed from all three hosts.
- Latest window-size verification/deployment: `artifacts\family-browser-ui-audit\20260711-103903-kky-tool-window-size` passed static/contract, five-target build/stage, staged verification, 2,000-row performance/cache, 112 IE DOM/click/language/layout/detail scenarios, ProgramData install, and installed verification. Revit 2021 has the expected one `SKIP runtime-not-installed`; failures are zero. Current ProgramData hashes match stage for all five targets.
- Latest window-size installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_kky-tool-window-size-20260713_Setup.exe`, 3.32 MB, SHA256 `0426F791F478F74877B6123DDE678F562200D00DBDE203BA376AE4E6C0F46C74`. Mail package: `artifacts\family-browser\mail-packages\20260713_01.zip`, 15.5 MB, SHA256 `5D46080164B8098257B606044FCAA8B71E786FD8FAE1953D44EED626825CFBED`. The embedded `Setup.exe` hash matches the standalone installer exactly.
- Latest window-size checkpoint: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260711\20260711-103704-kky-tool-window-size`.
- Latest language-transition fix: direct English rendering was already clean, but Korean -> English switching left cached Korean values in the standard summary, permission/tracking pills, unsaved-project placeholder, window title, and stored Family/System discipline labels. These values now re-localize from semantic keys/known states on every language switch, and the native form title updates immediately.
- Latest language QA fix: all English dashboard scenarios now initialize their cached display state in Korean before switching to English. The IE harness checks visible Hangul leakage, the translated unsaved-project subtitle, and the window-title diagnostic. The focused 2025 matrix passed 28/28; the full gate passed 112 `OK`, one Revit 2021 `SKIP runtime-not-installed`, and zero failures, including 56 English transition results with zero language failures.
- Latest deployment: 2019/2021/2023/2025/2027 were built, staged, installed, and verified in ProgramData. Stage and installed hashes match. Current installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_language-toggle-purity-20260711_Setup.exe`; mail package: `artifacts\family-browser\mail-packages\20260711_04.zip` (15.2 MB).
- Latest language-fix checkpoint: `artifacts\source-checkpoints\20260711-095147-language-toggle-purity`.
- Current design decision supersedes the earlier light/dark toggle work: Family Browser now has one fixed KKY Tool-aligned default palette. The header exposes only Refresh, Admin mode, and Language; there is no theme control and the UI does not label itself as light or dark.
- `FamilyBrowserUiThemeService` always resolves the fixed default palette, ignores any stale `theme.txt`, and no longer writes theme preference. The visible `kkyfb:theme-toggle` route/button/handler was removed from all three hosts and is now a forbidden contract token.
- Current automated coverage runs one palette in Korean/English while retaining brand-color, contrast, legacy-green, layout, detail, structured-message, click, route, and 2,000-row performance checks. The focused pass completed 10/10 and the full five-target gate completed 112 `OK`, one Revit 2021 `SKIP runtime-not-installed`, and zero failures.
- Current deployment: 2019/2021/2023/2025/2027 were built, staged, installed, and verified. Stage and ProgramData DLL SHA256 values match for all five targets. Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_light-only-20260711_Setup.exe`; mail package: `artifacts\family-browser\mail-packages\20260711_03.zip` (15.25 MB).
- Current checkpoint: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-checkpoints\20260711-020736-family-browser-light-only`.
- Latest result/error dialog fix: the shared `FamilyBrowserModernMessageDialog` and new `FamilyBrowserMessageHtmlRenderer` now live in `KKY_FamilyBrowser_SharedUi`, so 2019/2021/2023/2025/2027 use one implementation. The plain multiline `TextBox` body was replaced with an IE-compatible HTML `WebBrowser` body that parses friendly-error headings into separate failure reason, next action, administrator information, and technical detail cards. Log path/support code render as labeled rows, long technical text scrolls inside its own region, short notices remain compact, and long/structured messages expose `내용 복사` / `Copy details`.
- Latest choice safety fix: closing a Yes/No dialog through the custom `X` or `Esc` now returns `No` instead of the default affirmative result. Native OK/Yes/No footer behavior and the raw `MessageBox` emergency fallback remain intact.
- Latest automated coverage: `FamilyBrowserModernMessageDialog.BuildHtmlForAudit(...)` plus `structured-message-ko/en` harness scenarios render the actual message HTML in WinForms IE, require semantic headline/cause/action/admin/technical/log/support-code regions, reject English Hangul leakage, reject body horizontal overflow, and confirm that the HTML body does not contain accidental clickable controls. The full gate produced 89 results: 88 `OK`, one Revit 2021 `SKIP runtime-not-installed`, zero failures; all eight structured-message runtime scenarios passed with zero warnings.
- Latest deployment: `artifacts\family-browser-ui-audit\structured-message-dialog-20260710-quality-gate` passed static/contract, five-target build/stage, staged verification, 2,000-row performance/cache, full HTML/click harness, ProgramData install, and installed verification. Stage/installed DLL hash prefixes match: 2019/2021/2023 `E0A512DEFE22`, 2025 `B8DB19C7374A`, 2027 `5790392B58F2`.
- Latest backup: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\structured-message-dialog-20260710-155755`.
- Latest managed-data read audit: traced homepage bootstrap -> runtime managed policy -> versioned Registry/Snapshots/Thumbnails/Projects -> unversioned StandardLists/BrowserCacheV2 -> validated LocalAppData cache. `Test-FamilyBrowserManagedData.ps1` now checks the first reachable homepage candidate, shared policy, registration/snapshot/project references, V2 artifact lengths/SHA256/detail records, and project-state references without writing to the managed folder. It is available from the quality gate through `-ManagedDataAudit`.
- Latest managed-data fixes: startup preload now captures Revit project identity on the UI thread and applies the homepage bootstrap on the background preload before `LoadOrCreate`, so the first browser open cannot prepare rows from an empty/stale root. `InvalidatePreparedDashboardDataAfterManagedPolicyChange(...)` now clears startup, dashboard, row, and comparison caches for deferred bootstrap changes plus visible homepage path refresh, profile/URL changes, permission refresh, and field-test path application, so no manual route can fall back to the original preload object. Project latest-scan alias save/read no longer stops merely because `_workspaceRoot` is empty; it resolves the homepage-managed Projects folder and accepts a name alias only when the current project file stamp matches, preventing same-name project collisions.
- Live managed-path result: homepage bootstrap `2026.05.19-kcim-test` was reachable, but its `I:\30. 협력사 전용폴더\00. BIM_KCIM\02. 패밀리\TEST` and `D:\TEST` candidates were both unavailable on this PC. Therefore live shared policy/snapshot/project contents could not be inspected and the audit correctly reported `UNAVAILABLE`, not a code failure. `%LOCALAPPDATA%\KKY\FamilyBrowser\Cache\v2` contained zero real source manifests and zero row caches. Re-run the audit after mapping `I:` or publishing a reachable homepage candidate.
- Latest verification/deployment: a complete fixture passed, then a deliberately missing comparison report failed as expected. The full static/contract, five-target, performance, and 80-scenario UI gate passed before the final shared invalidation helper; the final helper then passed static/contract checks, a fresh five-target build/stage, staged verification, and the 2,000-row performance/cache gate. Current stage was installed and verified in ProgramData for 2019/2021/2023/2025/2027; final stage-vs-installed DLL SHA256 matched all five targets. Reports: `artifacts\family-browser-ui-audit\managed-data-read-20260710-complete-quality-gate` and `artifacts\family-browser-ui-audit\managed-data-manual-refresh-20260710-gate`.
- Latest backup: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\managed-data-read-audit-20260710-151626`.
- Latest incremental backup: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\managed-data-manual-refresh-invalidation-20260710-154314`.
- Latest setup-state fix: Family/System empty states now distinguish `RVT not registered` from `RVT connected + approved standard list not connected`. `IsActiveStandardListRegistrationRequired()` checks the selected policy slot before the registration-record fallback, so an RVT-connected target shows `등록된 표준 RVT의 표준 목록을 연결해주세요` and only the `표준 패밀리/타입 목록 등록하기` action. A genuinely unregistered target continues to show `표준 RVT를 먼저 등록해주세요` and the RVT registration action.
- Latest regression coverage: added `admin-family-missing-standard-list` beside the existing System scenario and `CheckStandardSetupEmptyState(...)`, which fails when the two CTAs/messages are mixed. The audit fixture also no longer equates missing RVT with scan-needed. A one-off 2019 IE queued-search timing failure was made stable by waiting up to 2.5 seconds for queued row rendering; the stable 2019 rerun passed 20/20.
- Latest deployment: current 2019/2021/2023/2025/2027 stage was installed to ProgramData and installed verification passed at 2026-07-10 14:57 KST. Stage-vs-installed DLL SHA256 matched for all five years. Revit 2021 runtime remains unavailable on this PC, but its build/stage/install package is verified.
- Latest backup: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\registered-rvt-missing-list-empty-state-20260710-142951`.
- Latest target-selection fix: Model Check `browse-discipline-*` changes now preserve the captured dashboard UI state instead of clearing it, so the current `mainCenter` / `auditPane` scroll position is restored after the cached shell refresh. Non-audit Family/System browse-target changes keep their intentional list/filter reset behavior.
- Latest Admin Standard layout fix: the selected trade chip now follows the actual current browse target in separated mode, so the displayed current target and pressed button cannot diverge. Trade add/rename/delete actions moved beside the Trade / library target selector. Baseline RVT and Visible Standard List actions now use equal-width/equal-height two-column rows plus full-width terminal actions.
- Latest regression coverage: the WinForms WebBrowser harness now performs an audit-scroll capture/reset/restore round trip, checks one active Admin trade target, verifies trade management is outside the Baseline RVT card, and compares action-row button widths/heights. Full report: `artifacts\family-browser-ui-audit\audit-scroll-admin-layout-20260710-complete-quality-gate`.
- Latest verification: 2019/2021/2023/2025/2027 build/stage, stage verification, static/contract checks, 2,000-row performance/cache gate, and 73 UI harness scenarios completed with 72 `OK`, one Revit 2021 `SKIP runtime-not-installed`, and zero failures. Revit 2019 remained open, so this pass did not update ProgramData.
- Latest backups: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\audit-target-scroll-preserve-20260710-134205` and `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\admin-target-actions-layout-20260710-140211`.
- Latest performance V2 implementation: the browser now publishes and consumes `family-browser-manifest-v2.json`, `standard-browser-index-v2.json`, `standard-browser-details-v2.json`, `thumbnail-index-v2.json`, `project-browser-state-v2.json`, and revision-keyed `browser-row-cache-v2-*` files. `%LOCALAPPDATA%\KKY\FamilyBrowser\Cache\v2` is the validated local read cache; shared/network data remains the source of truth, offline mode uses only the last validated cache, and model-mutating family load/system apply/tracking paths reject stale source revisions before execution.
- Latest startup/render implementation: the form renders a small startup shell first, loads policy/registration/compact snapshot/list data on a cancellable background generation, removes UI-thread `Task.Wait`, reuses same-key load work, performs partial active-pane replacement, hydrates selected detail/3D preview on demand, uses a one-time thumbnail index, and now sends Family/System rows as a compact positional JSON payload to `window.KKYFB.setRows(...)`. Only the active 150-row slice is created in the DOM; search/filter still cover all indexed rows.
- Latest DOM virtualization fix: the old all-row markup/hide-only window was replaced with `family-browser-row-window.js`, stable row keys, cross-page checked-item adapters, partial-pane payload refresh, saved-row page restoration, and duplicate-search/window render guards. Page 1 -> page 2 -> page 1 checkbox persistence and `Clear Selection` clearing both checks and selected-row state are enforced by the harness.
- Latest automated performance gate: `Test-FamilyBrowserPerformance.ps1` now records total, DOM, and visible rows separately and rejects hide-only virtualization. Final five-target runtime results were exactly 1,000 total / 150 DOM / 150 visible for both Family and System tabs. Shell was 1-12ms, cold usable 1,421-2,172ms, warm usable 245-347ms, changing-query filter response 9-10ms, and cache cold/warm/offline 16-68ms. All enforced targets passed.
- Previous scroll/layout baseline verification: `artifacts\family-browser-ui-audit\audit-scroll-admin-layout-20260710-complete-quality-gate` passed on 2026-07-10 14:17 KST. Static/contract checks found 77 generated actions, 244 exact routes, 50 prefix routes, and 11 browser-only functions in each host with no unknown route. Five-version build/stage, staged verification, 2,000-row performance, and Korean/English permission/registration/detail/layout/click harness passed with zero failures. Revit 2021 runtime smoke was `SKIP runtime-not-installed`.
- Previous DOM-virtualization deployment: ProgramData installation and installed verification passed for 2019/2021/2023/2025/2027 after the final installer rebuild. Stage-vs-installed DLL SHA256 matched for all five targets.
- Latest installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_dom-virtualization-20260710_Setup.exe`, SHA256 `FA190E3610B85BED52F2019958775655178BD44F67A1FFBDC86510E7D5F41D35`. Mail package `artifacts\family-browser\mail-packages\20260710_02.zip` is 14.8MB, SHA256 `E92FCF35CECD45C2E6CC407069345BDB82AF4FC6BD41C2EB37632A0116E2D26E`.
- Latest backups: DOM virtualization backup `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\family-browser-dom-virtualization-20260710-111520`; performance/cache source backup `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\family-browser-performance-v2-20260710-093316`; gate/harness backup `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\family-browser-performance-gate-20260710-101349`; startup-shell measurement/audit-MD backup `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\family-browser-performance-shell-audit-20260710-104904`.
- Latest deployment default fix: Family Browser build/quality-gate defaults now include Revit 2021 everywhere by default. `Invoke-FamilyBrowserQualityGate.ps1` and `Invoke-FamilyBrowserUiAuditHarness.ps1` now default to 2019/2021/2023/2025/2027, matching the existing build/install/verify/installer scripts. The installer output name and Inno file set already include `(2019,21,23,25,27)`.
- Latest harness fix: UI harness now records `SKIP runtime-not-installed` for a requested year when that Revit runtime is not installed on the current PC, while staged addin verification still checks the year-specific manifest/package. This prevents a missing local Revit 2021 install from blocking package generation. The search-focus harness also now derives a stable search token from row attributes instead of blindly using the first 10 characters of the row name.
- Latest verification: Quality gate passed on 2026-07-09 23:38 KST with target years 2019/2021/2023/2025/2027. Static/contract checks, Release build/stage, staged add-in verification, and WinForms WebBrowser UI harness passed; 2021 UI runtime smoke was explicitly skipped because `C:\Program Files\Autodesk\Revit 2021` is not installed on this PC. ProgramData installed verification passed for all five years after installing the 2021 addin folder. Revit 2019 was open during this work, so only the 2021 ProgramData folder was written in this step.
- Latest installer package: Family Browser installer for 2019/2021/2023/2025/2027 was rebuilt on 2026-07-09 23:45 KST. Installer exe is `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_include-2021-defaults-20260709-2345_Setup.exe`, SHA256 `7E646AD49C889E1B8B9EAB41C17298563C48EEE040982514174BE622B533F41F`. Mail zip is `artifacts\family-browser\mail-packages\20260709_04.zip`, 13.6 MB, SHA256 `D1F4837A0204EF125C8579FBC13C6D4A6E6C333C1DBF6C0E75EAFC3731F4F61E`.
- Latest backup: Script/test/audit-md backup before making 2021 part of the default build and quality-gate flow is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\include-2021-build-defaults-20260709-231700`.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt, staged, quality-gated, installed to ProgramData, and installed-verified on 2026-07-09 19:46 KST after the precise-scan dialog OK/Cancel/XLSX patch. Harness output: `artifacts\family-browser-ui-audit\scan-dialog-ok-cancel-xlsx-20260709-1700`; 72 scenarios passed with 0 failures. Revit was not running during install.
- Latest source fix: Standard RVT precise-scan dialog handling now records every auto-handled dialog with action and visible button data. Native dialogs are button-driven: if an enabled OK/Confirm/Continue button exists on a family-scan warning/error, the guard presses OK; if there is no OK and the active choices are Delete Instance/Delete Type/Delete Constraints/Remove Constraints plus Cancel, the guard presses Cancel. Revit `DialogBoxShowingEventArgs` cases still use override results, but destructive delete/remove text is routed to Cancel and generic scan warning/error text is routed to OK.
- Latest diagnostic export fix: Auto-handled standard scan dialog diagnostics now save as `.xlsx` instead of `.txt`, using the `ScanDialogs` sheet with `HandledAtUtc`, `Category`, `Family`, `Action`, `Reason`, `OverrideResult`, `AvailableButtons`, and `DialogText` columns. Result summaries now label this as `진단 Excel`.
- Latest verification detail: Static QA now guards the OK/Cancel action fields, Delete Instance/Delete Type detection, delete-only native cancel path, no native fake-IDOK fallback, `.xlsx` diagnostic path, and diagnostic Excel result label. Full quality gate and ProgramData installed verification passed for 2019/2023/2025/2027.
- Latest backup: Source/test/audit-md backup for the precise-scan dialog OK/Cancel/XLSX patch is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\scan-dialog-ok-cancel-xlsx-20260709-1700`.
- Latest source fix: Selected System Type apply no longer runs the full `SystemTypePreflightBuilderService.BuildReport(...)` path when multiple rows are checked. The selected path now builds and merges selected-only reports through `BuildSelectedSystemTypePreflightReport(...)` / `BuildSelectedReport(...)` for both pre-confirmation and post-apply verification, so already-scanned catalogs are not re-swept just because 2+ system types are selected.
- Latest UI/result fix: Family load and System Type apply confirmation/result dialogs now use `FamilyBrowserOperationHtmlDialog` HTML/WebBrowser content with summary cards, item result tables, notes, and diagnostic/report path actions instead of plain text WinForms message/result dialogs. The post-apply shell/dashboard refresh is wrapped in visible progress text, so the loading after the result window is identified as browser row/dashboard refresh.
- Latest verification: Full quality gate passed on 2026-07-09 16:16 KST for 2019/2023/2025/2027. Static/contract checks, Release build/stage, staged add-in verification, and WinForms WebBrowser click harness passed. Harness output: `artifacts\family-browser-ui-audit\system-apply-selected-html-results-20260709-1610`; 72 scenarios passed with 0 failures. This run did not install to ProgramData.
- Latest backup: Source/test/audit-md backup for the selected-system-apply and HTML operation result patch is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\system-apply-selected-html-results-20260709-155842`.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-09 15:53 KST after the system-type detail data patch. Static/contract checks, 2019/2023/2025/2027 Release build, staged add-in verification, full WinForms WebBrowser UI harness, ProgramData install, and installed verification passed. Harness output: `artifacts\family-browser-ui-audit\system-type-detail-data-20260709-1548`; 72 scenarios passed with 0 failures. Revit was not running during install, so restart/open Revit after this install to load the updated DLLs.
- Latest source fix: System Type detail rows now carry a captured detail summary instead of an empty parameter area. `SystemTypeDetailSummaryService` records identity, routing preference rules, segment/size counts, dependent loadable families, and layer rows into `@system-detail-v1`; semantic, standard, project, comparison, and preflight rows propagate that detail into `SystemRow.ParameterSummary`; the main panel and detached detail window render it as sectioned tables.
- Latest verification detail: The audit fixture now includes `AUDIT_DUCT_SEGMENT`, `sizes=12`, `AUDIT_DUCT_ELBOW`, and layer data for `AUDIT_SUPPLY_AIR`. Static QA guards the capture service, row propagation, renderer, fixture, and harness tokens. The WebBrowser harness selects the System Type row and fails if the system detail table, routing/segment, size count, dependency, or layer text is missing.
- Latest backup: Source/test/audit-md backup for the system-type detail data patch is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\system-type-detail-data-20260709-1508`.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-09 14:58 KST after the lookup CSV detail/difference patch. Static/contract checks, 2019/2023/2025/2027 Release build, staged add-in verification, full WinForms WebBrowser UI harness, ProgramData install, and installed verification passed. Harness output: `artifacts\family-browser-ui-audit\20260709-145141`; 72 scenarios passed with 0 failures. Revit was not running during install, so restart/open Revit after this install to load the updated DLLs.
- Latest source fix: Revit family internal imported CSV / lookup size tables were already included in the precise fingerprint through `lookup-tables=`. This patch makes that data visible and reviewable: detached/detail parameter preview now shows CSV presence and table size as `CSV 테이블` / `Lookup CSV` with row x column counts, and fingerprint difference details now split lookup CSV differences into table rows for project-only, standard-only, row/column-count mismatch, or content mismatch. Row/column/content mismatch still classifies the family as standard-different through the existing fingerprint difference path.
- Latest verification detail: The audit fixture now includes `AUDIT_SIZE_TABLE` in both detail preview and difference modal data. Static QA guards lookup table columns/rows capture, lookup CSV detail rendering, lookup CSV diff rendering, Korean labels, and harness checks. The WebBrowser harness fails if the detail table omits the CSV row/column summary or if the difference modal omits the lookup CSV row.
- Latest backup: Source/test/audit-md backup for the lookup CSV detail/difference patch is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\lookup-csv-detail-diff-20260709-143027`.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-09 14:25 KST after moving Debug Log to the left-menu-only / bottom-docked console behavior. Static/contract checks, 2019/2023/2025/2027 Release build, staged add-in verification, full WinForms WebBrowser UI harness, ProgramData install, and installed verification passed. Harness output: `artifacts\family-browser-ui-audit\debug-log-bottom-dock-20260709-1418`; 72 scenarios passed with 0 failures. Revit must be restarted/opened after this install to load the updated DLLs.
- Latest source fix: The floating/right-side `fbDebugFab` Debug Log button is no longer rendered. Debug access remains through the left navigation `디버그 로그` menu and F12. The debug panel now uses bottom-console docking: aligned after the left nav, attached to the viewport bottom, with fixed console height and no modal-style centered/top overlay. CSS asset overrides also hide any legacy FAB markup if it ever reappears.
- Latest verification detail: The UI harness now includes `CheckDebugDock(...)`, which fails if a floating debug FAB exists, if the debug panel is not attached to the bottom, if it sits too high like a modal, or if it overlaps the left menu. Static QA now guards the missing FAB anchor, bottom-dock inline CSS, CSS asset override, and harness failure messages.
- Latest backup: Source/CSS/test/audit-md backup for the debug-log bottom-dock patch is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\debug-log-bottom-dock-20260709-141215`.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-09 14:09 KST after fixing `LoadAvailable` family detail behavior. Static/contract checks, 2019/2023/2025/2027 Release build, staged add-in verification, full WinForms WebBrowser UI harness, ProgramData install, and installed verification passed. Harness output: `artifacts\family-browser-ui-audit\load-available-standard-detail-20260709-1402`; 72 scenarios passed with 0 failures. Revit must be restarted/opened after this install to load the updated DLLs.
- Latest source fix: Family rows with raw status `LoadAvailable` now render detached/detail content from the approved standard family snapshot even in Admin Mode. Because there is no current-project comparison target, the fingerprint/type difference table is suppressed, difference-style memo tails such as `타입 0/2` are not appended, nested children come from the standard snapshot, and the existing standard thumbnail resolution path remains in use. Non-`LoadAvailable` update/different rows still expose the concise fingerprint summary and `상세 보기` modal.
- Latest verification detail: The audit fixture now has a `LoadAvailable` row with standard parameters/nested children/preview and no diff table, plus a separate update row with fingerprint diff rows. The UI harness now fails if a load-available detail shows a fingerprint diff summary/button, and separately selects a diff row to verify the fingerprint diff modal still opens as a table.
- Latest backup: Source/test/audit-md backup for the load-available standard-detail patch is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\load-available-standard-detail-20260709-135746`.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-09 13:52 KST after fixing live search focus loss in the Family/System Type lists. Static/contract checks, 2019/2023/2025/2027 Release build, staged add-in verification, full WinForms WebBrowser UI harness, ProgramData install, and installed verification passed. Harness output: `artifacts\family-browser-ui-audit\search-focus-silent-detail-20260709-1345`; 72 scenarios passed with 0 failures. Revit must be restarted/opened after this install to load the updated DLLs.
- Latest source fix: Family/System Type live search now uses a search-only quiet detail update path. `queueFilterRows()` calls `filterRows('search')`, `filterRows` selects the first visible row with `quietDetail`, and the wrapped `selectRow` temporarily suppresses `detail-window-sync`, `detail-window-open`, and `preview-inline/*` host actions while preserving the search input focus. Normal row clicks, tab-entry auto detached detail open, toolbar `상세 항목`, and manual preview behavior are unchanged.
- Latest verification detail: The UI harness now includes `CheckSearchFilteringKeepsFocus(...)`, which focuses `searchBox`, runs queued filtering in Family/System scenarios, verifies the detail panel follows the filtered first row, and fails if search filtering emits focus-stealing host actions. Static QA now guards the quiet-search tokens in all 3 host files and the harness failure messages.
- Latest backup: Source/test/audit-md backup for the live-search quiet-detail patch is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\search-focus-silent-detail-20260709-132803`.
- Latest installer package: Family Browser installer for 2019/2021/2023/2025/2027 was rebuilt on 2026-07-09 13:16 KST after the Model Check review-target selector patch. Installer exe is `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_model-check-target-20260709-1021_Setup.exe`, SHA256 `394C831D072E3F32C38A00D1E4DB64DA43745A7340D61097C72E69E86C07CAB9`. Mail zip is `artifacts\family-browser\mail-packages\20260709_03.zip`, 13.6 MB, SHA256 `F47DC01392CADEADF309DA5C686831E9BA8B46DCCD7AA9B31292BAE2AE96A0F6`.
- Latest source fix: Model Check now has an inline `audit-target-selector` under `선택된 검사 기준`, so the user can switch the review/standard target directly from the Model Check tab instead of going to Standard Management first. The chips reuse the existing `kkyfb:browse-discipline-*` route and `SetBrowseDiscipline(...)`, show the active target, mark unscanned targets as `스캔 필요`, and use a Model Check-specific status message after switching. While validating this, the Model Check card grid was also hardened with IE WebBrowser-safe float-based 2x2 layout so the four action cards stay two per row until the actual content width is narrow.
- Latest verification: Static/contract checks passed, 2019/2023/2025/2027 Release build passed, staged add-in verification passed, and the WinForms WebBrowser harness passed on 2026-07-09 10:21 KST. Harness output: `artifacts\family-browser-ui-audit\audit-target-selector-20260709-1015`; 72 scenarios passed with 0 failures.
- Latest install: 2019/2023/2025/2027 ProgramData install and installed verification passed on 2026-07-09 10:21 KST. Revit was not running during install. Revit must be restarted/opened after this install to load the updated DLLs.
- Latest backup: Source/CSS/test backup for the Model Check target selector and related layout/harness patch is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\audit-target-selector-20260709-094538`.
- Latest source fix: Admin Standard Management action cards now use scoped `admin-standard-action-grid` markup and IE WebBrowser-safe inline-block two-column layout, so `기준 RVT` and `표시 표준 목록` sit side by side with compact wrapping buttons instead of full-width stacked rows. Model Check cards now use scoped `audit-action-grid` markup, rendering the four action cards as a 2x2 layout while trade/action buttons wrap inside the card instead of widening or clipping. The UI harness now includes `admin-model-check-layout` and checks both action-card grids for side-by-side rows and button overflow.
- Latest verification: Static/contract checks passed, 2019/2023/2025/2027 Release build passed, staged add-in verification passed, and the WinForms WebBrowser harness passed on 2026-07-09 09:42 KST. Harness output: `artifacts\family-browser-ui-audit\admin-audit-card-grid-20260709-0936`; 72 scenarios passed with 0 failures.
- Latest install: 2023/2025/2027 ProgramData install and installed verification passed on 2026-07-09 09:43 KST. Revit 2019 was open (`Revit.exe` PID 15504, `Autodesk Revit 2019.2.6 - [Project1 - Floor Plan: Level 1]`), so the 2019 ProgramData install was intentionally left pending until Revit 2019 is closed.
- Latest backup: Source/CSS/test backup for the Admin Standard / Model Check card-grid patch is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\admin-audit-card-grid-20260709-092656`.
- Latest automation fix: Family Browser UI harness now runs every contract scenario in both Korean and English language variants, suffixing result scenario names with `-ko` / `-en`. The harness adds `CheckLanguagePurity(...)`: English mode fails if visible Korean text remains, except the intentional `한국어` language-switch label; Korean mode fails if core Korean UI tokens are missing or common untranslated English UI phrases appear outside the allowlist for technical/product terms such as KKY, RVT, Revit, Excel, CSV/JSON/PDF/PNG, F12, ON/OFF, file paths, and audit fixture data. Static QA now guards the language check call, English/Korean failure messages, disallowed-English list, and wrapper language-matrix execution.
- Latest verification: Static/contract checks passed, UI harness project built with 0 warnings/errors, and the full language matrix harness passed on 2026-07-09 09:23 KST for 2019/2023/2025/2027. Harness output: `artifacts\family-browser-ui-audit\language-purity-20260709-0914`; 64 scenarios passed with 0 failures.
- Latest backup: Source/test backup for language purity automation is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\language-purity-audit-20260709-091434`.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-09 09:12 KST after the main header file-context layout patch. Static/contract checks, build, staged verification, WebBrowser UI harness, ProgramData install, installed verification, and stage-vs-installed SHA256 comparison passed. Installed hash prefixes: 2019/2023 `74430138D4F1`, 2025 `34267BD960AB`, 2027 `1D8960F086B6`. Revit must be restarted after this install to load the updated DLLs.
- Latest source fix: Main dashboard header no longer renders the current project title as a duplicate file token. Local and central project files now render as stacked right-aligned rows in `project-context`, with full paths retained in hover titles. `새로고침`, Admin Mode, and language toggle are separated into top-right `top-actions`, while status pills render in a lower `status-pills` row with extra vertical breathing room. The shell CSS now pushes the layout/workflow body below the enlarged header. Static QA and the WinForms WebBrowser harness now check the separated top actions, stacked local/central rows, missing duplicate project title token, and spacing between actions, file context, and status pills.
- Latest backup: Source/CSS/test backup for the header patch is `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\header-file-context-layout-20260709-090225`.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-08 17:23 KST after fixing duplicate composite nested-child rows with dash categories. Static/contract checks, build, staged verification, WebBrowser UI harness, ProgramData install, and installed verification passed. Revit must be restarted after this install to load the updated DLLs.
- Latest source fix: Composite/nested family detail now dedupes child rows by nested family name. If the same child appears once with category `-` and once with a real category, `BuildNestedLoadableSummary` and the detached/detail JS parser keep the real categorized row and drop the dash row. The audit fixture now deliberately injects duplicate dash/category child rows, and the WebBrowser harness fails if the detail table keeps a dash duplicate or renders more than two child rows for that fixture. The harness wrapper also has a Windows PowerShell-compatible `ProcessStartInfo.Arguments` fallback when `ArgumentList` is unavailable.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-09 08:56 KST after fixing the detached detail parameter value/formula table. Static/contract checks, build, staged verification, WebBrowser UI harness, ProgramData install, and installed verification passed. Revit must be restarted after this install to load the updated DLLs.
- Latest source fix: Detached detail parameter tables now render through an internal `parameter-table-scroll detached-parameter-scroll` wrapper. The non-formula columns use controlled widths close to the existing instance-parameter table proportions (`kind/scope/name/value`), the formula column gets a wider dedicated column, and long formulas stay inside the parameter area with horizontal table scrolling instead of pushing the detail page outside its bounds. Added a long formula audit fixture and static guards for the scroll wrapper, table width, formula column, and override renderer.
- Latest temp-execution workaround installer: Standard Inno Setup installers can still launch an internal setup loader from `%TEMP%`, and running an `.exe` from inside a mail zip / attachment preview can also execute from `%TEMP%`. On PCs that block installer execution from temp, this can surface as a "cannot install from temporary directory" error even when the visible file appears to be on Desktop. A `UseSetupLdr=no` installer was compiled on 2026-07-09 08:24 KST to avoid the Inno setup-loader temp extraction path. Installer exe is `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_nested-child-dedupe-20260709-0824_NoSetupLdr_Setup.exe`, SHA256 `64F69F9996115532869B3B4F9685DCAA57584F31F718A6B717653FB3074F8059`. Users should extract/copy it to a real folder such as `C:\KKYInstall` or Desktop and run it as administrator, not launch it from inside a zip or mail preview.
- Latest installer package: Family Browser installer for 2019/2021/2023/2025/2027 was rebuilt again on 2026-07-09 08:19 KST from the same nested-child-dedupe source. Installer exe is `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_nested-child-dedupe-20260709-0819-rebuild_Setup.exe`, SHA256 `B6D9F8E915C44131F3D55C97E089C737B5EF691C813B7765B431C368D0D1A14A`. Mail zip is `artifacts\family-browser\mail-packages\20260709_02.zip`, 13.6 MB, SHA256 `28A78C7CC9CED299FD88F2D8EAA92188FFDB1DD345C454BF990C254D743A0EEE`.
- Latest installer package: Family Browser installer for 2019/2021/2023/2025/2027 was rebuilt on 2026-07-09 08:16 KST from the nested-child-dedupe source. Installer exe is `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_nested-child-dedupe-20260709-0816_Setup.exe`, SHA256 `BD0722E76C96EDC1A376DDD02359D41AB6D8C3B9AA4CC0531C9D2C1D9F1AA80C`. Mail zip is `artifacts\family-browser\mail-packages\20260709_01.zip`, 13.6 MB, SHA256 `52E16E59865C07D9EBDB66EDA764A45B548C91F7B11B55FCF62A3BECCBD62C16`.
- Previous installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-08 17:08 KST after fixing detached detail family composition/type-parameter/3D preview layout. Static/contract checks, build, staged verification, WebBrowser UI harness, ProgramData install, and installed verification passed.
- Previous source fix: Detached detail `패밀리 구성` now renders family type lists inside a styled `family-type-table` panel instead of visually reading as loose `# / 타입 이름` text. Nested/composite family rows still use the category/family table. Type-parameter dropdown spacing and width were increased so long type names have more room and the native dropdown is not cramped inside the panel. Small 3D preview fit no longer depends on percent-position plus CSS transform; it computes pixel left/top center positions and the detached preview cell now uses IE-compatible left/right/top/bottom bounds instead of only `inset:0`. Added static guards for the type table, dropdown width, and non-transform preview centering.
- Latest installer package: Family Browser installer for 2019/2021/2023/2025/2027 was rebuilt on 2026-07-08 16:31 KST after the second Family/System Type action-row status filter wrap patch. Installer exe is `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_inline-status-wrap-20260708-1630_Setup.exe`, SHA256 `09AAEDA170BA722CC0C450B3D47961335AAB3221723EF57CF73133E943C2F37F`. Mail zip is `artifacts\family-browser\mail-packages\20260708_03.zip`, 13.6 MB, SHA256 `4E48B732E951EB74B7146DB6E94BAE01930780A980D493EAA82CE40B32DBCF40`.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-08 16:30 KST after fixing the second Family/System Type action-row status filters. Static/contract checks, build, staged verification, WebBrowser UI harness, and installed verification passed. Revit must be restarted after this install to load the updated DLLs.
- Latest source fix: The search-area status/trade filter bars had already been fixed, but the action row beside `선택 항목 로드` / `선택 초기화` still used `.inline-status-toggle` with fixed max width, hidden overflow, and ellipsis. The shell JS also reapplied a fixed 58/56 px header height and hidden inline status width at runtime, so CSS-only changes could be overwritten inside Revit. Backed up sources to `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\inline-status-filter-wrap-20260708-1610`, changed all 3 host projects so the action row wraps with IE-compatible flex, measures the real wrapped header height, and moves the family/system grid below that measured height. Added dense audit rows for matched/name-check/update/different/review statuses and extended the WebBrowser harness to fail if inline action status filters use ellipsis, clip text, or overlap the grid.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-08 16:05 KST after fixing detached/detail fingerprint difference display. Static/contract checks, build, staged verification, WebBrowser UI harness, and installed verification passed. Revit must be restarted after this install to load the updated DLLs.
- Latest source fix: Detail `기준 상태` now keeps fingerprint mismatch text concise (for example `타입 수 다름`) and moves the detailed evidence behind a stronger `상세 보기` button. The button now carries encoded raw diff rows (`data-diffraw`) so it works both in the main browser and copied detached detail window, instead of depending only on an in-page store id. Detached detail HTML now includes the same fingerprint diff parser/table renderer and a real diff modal; the old `openDiffModal(){return false;}` no-op is guarded against by static QA. The audit fixture now includes Type Count/Parameter diff rows, and the WebBrowser harness clicks `상세 보기` and verifies the `fingerprint-diff-table` modal content.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-08 15:38 KST after the Admin Settings/detail preview layout patch. Static/contract checks, staged verification, UI harness, and installed verification passed. Revit must be restarted after this install to load the updated DLLs.
- Latest source fix: Admin Settings `표준 기준 / 표시 목록` layout was changed from cramped mixed inline cards to a compact two-column grid: status/standard selectors stay on the left and the action cards sit in a fixed-width right column that stacks only on narrow windows. Detached detail title now uses the generic selected-item title so the family name appears only once in the detail hero with its category chip. Detached detail parameters stay full-width, complex loadable families render nested child families as a category/family table, and the 3D preview `크게 보기` chip is a real clickable modal action. Audit fixtures now include a composite family and the WebBrowser harness checks nested table content, 3D preview markup, and large-view modal opening. Existing scans only need to be rerun when no PNG thumbnail cache exists; if the PNG exists but Revit WebBrowser cannot display the file URI, the detail window uses the inline data-URI fallback.
- Latest installer package: Family Browser installer for 2019/2021/2023/2025/2027 was rebuilt on 2026-07-08 14:55 KST after the Family/System Type detached-detail auto-open and detached-detail layout patch. Installer exe is `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_detail-auto-open-20260708-1455_Setup.exe`, SHA256 `4618EC08CC73729211A26E2FCA73AE8015CB5B55F578FB04FF1416304D5ED677`. Mail zip is `artifacts\family-browser\mail-packages\20260708_02.zip`, 13.6 MB, SHA256 `A79F93B2DE515D66EFDC8727233EDE3612B9BC854ABD2ACA2264DA30B6B868F0`.
- Latest installed build: Family Browser 2019/2023/2025/2027 was installed to ProgramData on 2026-07-08 14:50 KST after the Family/System Type detached-detail auto-open and detached-detail layout patch. Static/contract checks, staged verification, UI harness, and installed verification passed. Revit must be restarted after this install to load the updated DLLs.
- Latest source fix: Entering the Family Load or System Type list now automatically opens the detached detail window when the current tab has visible rows. The logic selects the current row or first visible row, emits `kkyfb:detail-window-open`, and keeps normal row clicks as selection/sync behavior. The detached detail window was widened and its internal layout was cleaned so identity/status, parameter/type data, and preview areas sit in stable sections. Added static guards, a WebBrowser harness auto-open assertion, and an `admin-system-with-data` audit scenario so both Family and System Type tab-entry behavior are checked. Static/contract checks, build, stage verification, WebBrowser harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027.
- Previous installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-08 14:33 KST after the project local/central subtitle tooltip patch. Static/contract checks, staged verification, UI harness, and installed verification passed.
- Latest source fix: the header subtitle under `KKY 패밀리 브라우저 1.0` now shows compact project context instead of a long inline path. Local/workshared files render as file-name tokens with the full local path in the hover title, and workshared documents also show the central file-name token with the central path in the hover title when it can be resolved. Added Revit-free audit scenario local/central paths, static regression guards, and WinForms WebBrowser harness checks for visible local/central names plus tooltip path titles. Static/contract checks, build, stage verification, UI harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027.
- Latest source fix: Family/System Type status filters and discipline filters could be clipped after a standard RVT scan because `.filterbar` / `.disciplinebar` were forced to one line with `white-space: nowrap`, `overflow: hidden`, `max-width`, and `text-overflow: ellipsis`. Updated all 3 host projects and shell CSS assets so these filter bars wrap with flex, buttons have no max-width cap, and labels are not ellipsized. Added static regression guards and extended the WinForms WebBrowser layout harness so rendered filters fail the audit if they use ellipsis or clip visible text. Static/contract checks, build, staged verification, UI harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027.
- Latest source fix: Admin Settings `기준 RVT` and `표시 표준 목록` button groups were reorganized on 2026-07-08 so primary actions are grouped by workflow instead of flowing as mixed inline buttons. `기준 RVT` now shows RVT connect / manage scan as the primary row, reset as a secondary row, and trade add/rename/delete in a compact trade-management row. `표시 표준 목록` now shows Excel connect as the primary row, then Fingerprint audit / RVT list creation, then blank template / clear list. Added `standard-action-*` CSS and static guards, plus a new `admin-standard-settings-layout` harness scenario. Static/contract checks, build, stage verification, and UI harness passed for 2019/2023/2025/2027. This fix is included in the 2026-07-08 14:20 all-version ProgramData install.
- Latest source fix: Standard-family 3D preview capture was traced from precise standard RVT registration through `FamilyThumbnailPreviewService.GenerateFromOpenFamilyDocument` and row `PreviewImagePath` / `data-previewpath`. The detached detail window depended on file/URI rendering and did not have the same inline data-URI fallback as the main panel, so captured PNGs could exist but fail to display in the separate detail window. Added detached-detail inline PNG fallback in all 3 host projects, added an audit fixture preview image plus realistic family/type parameter summaries to `FamilyBrowserDashboardAuditScenario`, and extended the WinForms WebBrowser harness to select an audit family row and verify detail name, category, type/composition, parameter values, and rendered 3D preview DOM. Static/contract checks, build, stage verification, UI harness, install, installed verification, and hash comparison passed for 2019/2023/2025/2027 on 2026-07-08 13:45 KST.
- Latest mail package: Family Browser installer for 2019/2021/2023/2025/2027 was rebuilt and packaged on 2026-07-08 13:17 KST. Mail zip is `artifacts\family-browser\mail-packages\KKY_FamilyBrowser_RevitHost_2019_2021_2023_2025_2027_v1.0_mail_20260708.zip`, 13.6 MB, SHA256 `E3FE50B4EBAEDFB75F530FCADD4272530495C68E37FAC57E619E3290ABF40A2C`. Installer exe inside the zip is 3.19 MB, SHA256 `BD89AF18F3578B8649FF66999206F5601EB888037CE4034825B3E69FAE7E98F5`.
- Latest source fix: Family Browser automated UI quality gate was implemented and stabilized on 2026-07-08. Added `FamilyBrowserUiAudit.contract.json`, static/contract checks, Revit-free audit HTML render seams, WinForms WebBrowser click harness, and `Invoke-FamilyBrowserQualityGate.ps1`. During harness validation, 2019/2023 generated dashboard JS had a real escaped-newline bug in `renderCompositionDetail` (`replace(/\r?\n/g...)` emitted as a broken regex and made browser-only buttons undefined); fixed it in `KKY_FamilyBrowser_RevitHost_2019-2023\FamilyBrowserDashboardHtmlForm.cs` and added a static regression guard. The `AnavRes.dll not found` popups were not a Family Browser runtime failure; they were caused by the external audit harness loading Autodesk/Revit native UI DLLs without the required Autodesk Shared RealDWG/Components paths. The harness now suppresses native DLL popups, hard-exits cleanly, and adds Revit language, RealDWG, and Components dependency directories. Integrated quality gate passed for 2019/2023/2025/2027 on 2026-07-08 11:41 KST.
- Latest source fix: Debug Log / F12 did not always show because the debug panel render gate used `showAdminUi && canViewAdmin` while the host route used `CanSeeInternalPaths()`, and dashboard JavaScript could still catch F12 when `#fbDebug` was not rendered. Aligned the render gate to `CanSeeInternalPaths()`, added host-side `TryToggleDashboardDebugConsole()` so the host verifies the panel exists before treating the toggle as successful, changed JavaScript so F12 first toggles an existing panel and otherwise routes `kkyfb:debug-log` to the host for the admin-only/page-not-ready message, and extended static QA. Built, staged, installed, installed-verified, and stage-vs-installed hash-checked 2019/2023/2025/2027 on 2026-07-06 16:28 KST.
- Latest installed build: Family Browser 2019/2023/2025/2027 was rebuilt and installed to ProgramData on 2026-07-06 16:27 KST; installed verification and stage-vs-installed SHA256 comparison passed on 2026-07-06 16:28 KST.
- Latest recheck: selected Family tab load path was re-traced on 2026-07-06 from UI JS `loadSelectedFamily` / `loadCheckedFamilies` through `sync-family/*` / `sync-families-selected/*`, `ApplySelectedStandardFamilyFromAction`, `ApplyStandardFamilies`, `BuildSelectedLoadableFamilyLoadPlan`, `StandardLibraryDocumentResolver.OpenRegisteredDocument`, and `LoadableFamilySyncExecutionService.Execute` to `familyDoc.LoadFamily(targetDocument, loadOptions)`. No source change was made. Code-level result: missing standard families in `LoadAvailable` state can be planned as `ExecutionMode = Load`; already-loaded/update rows are intentionally skipped in the normal Family tab and remain Admin Settings existing-family update flow. Revit runtime check is still required for an actual standard RVT/current model load.
- Latest source fix: browser-button navigation recheck found that the main dashboard already accepted `about:kkyfb:`, but nested WebBrowser surfaces for request drafting, standard-family update selection, the Standard RVT manager, file-guard configuration, and standard-list sheet selection still depended on only the direct custom scheme form. Added `about:` scheme guards for `kkyfb-request:`, `kkyfb-select:`, `kkyfb:`, `kkyfileguard://`, and `kkysheet://` in all 3 host projects, and extended static QA so these routes cannot silently regress. Static action-route extraction found 0 unknown dashboard actions after the patch.
- Latest source fix: dashboard, standard-family selection, and file-specific guard confirmation dialogs now route through shared `FamilyBrowserModernMessageDialog` in all 3 host projects. Direct raw `MessageBox.Show` calls were removed from those UI paths, `ShowDashboardMessage` / `ShowDashboardChoiceMessage` now use the shared modern dialog route, and static QA now guards against reverting those routes. Revit external command / native command guard `TaskDialog.Show` calls remain as Revit-native command notifications and are a separate `Needs Design` cleanup if the user wants every ribbon/native-command result surface replaced.
- Latest source fix: Current Model Check result export is now intentionally `.xlsx` Excel workbook only. The previously added review-result `.csv` export was based on a misunderstanding and has been removed from the save dialog, export service, and static QA. The result dialog still keeps the `Excel 추출` button, and the service normalizes any non-`.xlsx` path to `.xlsx` before writing the workbook.
- Latest source fix: Revit family internal imported CSV / lookup size tables are now explicitly protected in the precise loadable-family fingerprint path. User clarified that the intended CSV was not the Current Model Check export CSV, but the CSV lookup table used inside `.rfa` families for size/type definitions. The signature path already had `lookup-tables=`, but `ResolveFamilySizeTableManager` only looked for a one-argument API even though Revit 2019/2023/2025 expose `GetFamilySizeTableManager(Document, ElementId)`. Added owner-family-id resolution and the two-argument manager call in all 3 host projects, so lookup table names, column headers, and `AsValueString` row/cell values can affect the family fingerprint. Revit runtime check still required with an actual lookup CSV family.
- Previous source fix corrected: Current Model Check review export keeps the difference columns for both loadable families and system types, but no longer exposes or writes `.csv`. Before this correction, loadable-family difference fields were exported, but system-type difference fields were blank; that useful `.xlsx` difference-column fix remains.
- Latest source fix: language/result-window QA pass completed after user reported repeated manual findings around result dialogs, garbled/clipped text, and Korean/English switching. Patched result dialog sizing/wrapping and deferred detached-detail refresh after language-render completion. Added and passed `KKY_FamilyBrowser_Compile\Test-FamilyBrowserUiStatic.ps1` so these checks run before future builds.
- Latest source fix: Permissions / Guard tab is now file-guard-only. Role lists, permission Excel controls, project rules, current permission check, and the O/X diagnostic table are no longer rendered there.
- Latest source fix: Family/System Type list command buttons were re-stabilized after detached-detail work. Row click now only selects rows, detached detail opens only from the `상세 항목` toolbar button, and silent load/apply permission returns now show a message.
- Previous source fix: Family/System Type `상세 항목` opens by accepting `about:kkyfb:` navigation and scheduling the detail-window action before preview/navigation side effects can swallow it.
- Latest source fix: when a precise standard RVT scan exists but the approved standard family/type Excel or JSON list is missing, Admin Mode now shows a `표준 패밀리 등록하기` action in Family/System Type empty states and routes it to the Admin Settings standard-list registration area.
- Latest source fix: governance/admin operational controls are exposed in the intended split tabs, not inside the Standard Admin Settings tab. Request store controls are on Requests, the Permissions / Guard tab is file-guard-only, and deployment/path/reset controls are on Help / Technical Details.
- User correction recorded: Admin Settings was intentionally split by tabs. The attempted Admin Settings consolidation was reverted to backup-equivalent text before the tab-scoped fix.
- Representative action inventory was re-scanned from `KKY_FamilyBrowser_RevitHost_2019-2023\FamilyBrowserDashboardHtmlForm.cs`; visible governance `kkyfb:` actions are represented in this table.
- Revit runtime check recommended next: open Family Browser in Admin Mode and verify Requests controls, the file-guard-only Permissions / Guard tab, and Help / Technical Details controls; also run/current-model-check enough to produce unregistered rows, then click the row-level `요청` action for both family and system type rows and confirm the request composer is prefilled correctly.

## Button Audit Table

| Area | Action / Button | Router / Handler | Intended Behavior | Finding | Fix / Result | Status |
|---|---|---|---|---|---|---|
| Permissions / Guard / Model Check | Per-RVT trade assignment, XLSX bulk policy, and first-open automatic Current Model Check | `FileGuardHtmlConfigurationForm` -> `FamilyBrowserFileGuardDisciplineService` / `FamilyBrowserFileGuardExcelService` -> `FamilyBrowserStandardPolicyStore.SetFileGuardPolicy`; `App.DocumentOpened` / `App.Idling` -> `FamilyBrowserAutomaticModelCheckService`; nested guard -> `FamilyBrowserNestedOnlyPlacementRuntimeService` | Each registered project RVT must identify exactly which trade standard governs it. The assignment must survive XLSX export/import, the first open must automatically restore or create that trade's Current Model Check result, and nested-only review must use the same trade only. | File Guard targets had no trade field, bulk Excel had no trade column/import path, Current Model Check remained manual, and nested-only candidate loading could aggregate every separated standard. This could compare a guarded RVT against the wrong trade or make the user switch standards manually before every project check. | Added persisted `Discipline`, per-row/bulk selector, Korean/English XLSX round-trip, exact registered-slot validation, first-open idle scheduling, cache reuse, project scan lock, comparison persistence, dashboard status, and assigned-trade nested-only cache isolation. Static/contract, `984/984` workflow, dedicated nested/system checks, 5-target build/Stage/install verification, 36 focused IE scenarios, and a real Korean XLSX export/import harness passed. Actual Revit API capture and two-session locking remain queue item 43. | Needs Revit Check |
| System Types | Detail data display | `SystemTypeSemanticCaptureService` / `SystemTypeDetailSummaryService` -> snapshot `DetailSummary` -> `SystemRow.ParameterSummary` -> `renderSystemDetailTable` / detached detail renderer | System Type detail should show enough captured standard data to understand the basis: identity, routing rules, segment/size counts, dependent loadable families, and layer composition where available. | User reported the System Type detail page had a `시스템 타입 검토 데이터` area but it was effectively empty. Code confirmed semantic scan already captured much of the routing/dependency fingerprint, but the UI row model did not carry a display detail payload and the detail JS hid parameter/detail content outside the Family tab. | Backed up sources, added `SystemTypeDetailSummaryService` in all 3 host projects, propagated `DetailSummary` through semantic/standard/project/comparison/preflight rows, rendered `@system-detail-v1` as sectioned tables in main and detached detail windows, added audit fixture data for `AUDIT_SUPPLY_AIR`, and extended static QA plus WebBrowser harness to verify routing/segment size/dependency/layer detail. Static/contract checks, build, stage verification, 72-scenario UI harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027. Real Revit routing content still needs a runtime check with an actual scanned standard RVT. | Fixed |
| Families / System Types | Live search focus and detail sync | Search input `queueFilterRows` -> `filterRows('search')` -> quiet `selectRow(first,false,quietDetail)`; harness `CheckSearchFilteringKeepsFocus` | Typing in the Family/System Type search box should keep the search field active while the list and detail panel update live. Detached detail windows and inline preview host calls should not steal focus on every typed character. | User reported that each Korean character typed in the search box caused the detail item/window to activate and the search box to lose focus. Code confirmed live filtering called `filterRows()`, which immediately reselected the first visible row through the normal `selectRow` path. That path can emit `detail-window-sync` and `preview-inline/*`, so the WinForms/Revit host could activate the detached detail window during typing. | Backed up sources, added a search-only quiet detail path in all 3 host projects, temporarily suppressing detached-detail sync/open and inline-preview host actions only during queued search filtering while keeping normal row clicks, tab-entry auto-open, toolbar detail opening, and manual preview behavior intact. Added static guards and a WebBrowser harness check that focuses `searchBox`, runs live search, verifies the detail follows the filtered first row, and fails if search emits focus-stealing host actions. Static/contract checks, build, stage verification, 72-scenario UI harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Model Check | Review target selector | `AppendCheckMaintenanceWorkspace` -> `AppendAuditTargetSelector` -> `kkyfb:browse-discipline-*` -> `SetBrowseDiscipline` | The selected review/standard target should be changeable directly on the Model Check tab, and later compare/update action cards should use that target without requiring a Standard Management round trip. | The tab only displayed `선택된 검사 기준` as read-only text from the current Standard Management/browser target. This forced users to switch tabs just to change the basis before running model checks. During validation the audit action-card grid also exposed an IE WebBrowser layout weakness: the harness viewport could render the four Model Check cards as a single column even when the content width had room for two columns. | Backed up sources, added the `audit-target-selector` chip group in all 3 host projects, reused existing target-switch routing, added active/missing target labels, added Model Check-specific status text, added CSS for the selector, fixed the Model Check cards with float-based two-column layout, and extended the WebBrowser harness/static contract to require the selector, active chip, route, and grid presence. Static/contract checks, 2019/2023/2025/2027 build, staged verification, 72-scenario UI harness, ProgramData install, and installed verification passed. | Fixed |
| Admin Settings / Model Check | Standard source/list card layout and model-check action cards | `AppendAdminPane`; `AppendCheckMaintenanceWorkspace`; shell CSS `.admin-standard-action-grid` / `.audit-action-grid`; harness `CheckLayout` | `기준 RVT` and `표시 표준 목록` should sit side by side with compact buttons. The four Model Check cards should render two per row, and extra trade/action buttons should wrap within the card instead of stretching right or clipping. | User screenshots showed the Standard Management action cards/buttons were too wide and stacked vertically, and Model Check cards used only the left half of the page while the right side stayed empty. Previous CSS relied on mixed `grid`/full-width overrides that are brittle in Revit's IE WebBrowser. | Backed up sources, added scoped action-grid classes in all 3 host projects, added final IE-safe inline-block CSS overrides in all 3 shell CSS assets, compacted button sizing/wrapping, added an `admin-model-check-layout` audit scenario, and extended the WebBrowser harness to fail if the action cards are not side-by-side or buttons escape their cards. Static/contract checks, 2019/2023/2025/2027 build, staged verification, and 72-scenario UI harness passed. 2023/2025/2027 were installed and verified; 2019 install is pending because Revit 2019 is open. | Fixed |
| Detached Detail | Parameter value/formula table width and overflow | `BuildDetachedSelectionDetailHtml`; detached JS `renderDetachedParameterRowsTable` / `renderParameterFormulaDetailTable` / `renderTypeParameterRowsHtml`; CSS `.detached-parameter-scroll` | The detached detail parameter area should keep kind/scope/name/value widths readable and stable, while long formula text should stay inside the parameter block and be reachable by scrolling the table horizontally. | User reported that the value/formula table columns were too uneven and long formulas could push the table outside the detail page to the right. Code confirmed the detached-detail renderer still emitted some parameter tables directly without the same internal scroll wrapper used elsewhere. | Backed up sources, added detached-only table scroll wrapper and width overrides in all 3 host projects, re-overrode detached parameter table renderers to always emit the scroll wrapper and formula cell class, added a long formula audit fixture, and extended static QA guards. Static/contract checks, build, stage verification, WebBrowser harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Detached Detail | Composite nested child duplicate rows | `BuildNestedLoadableSummary`; JS `addNestedChildRow` / `parseNestedChildRows`; harness `CheckBrowserDetailContent` | Complex family composition should list each child family once. If one detected row has category `-` and another row has the real category for the same child family, the real category row should be kept and the dash row should be removed. | User reported composite-family child families appeared twice, with one duplicate carrying category `-`. Code confirmed both the summary builder and detail parser could treat `-\tFamily` / `Family` and `Category\tFamily` as different rows because the category was part of the uniqueness decision. | Backed up sources, changed the summary builder to dedupe by family name and prefer non-dash category rows, overrode the detail parser with the same rule, added duplicate dash/category audit fixture rows, and extended static QA + WebBrowser harness to fail if dash duplicates remain or row count is not deduped. Static/contract checks, build, stage verification, UI harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Detached Detail | Family composition, type parameter dropdown, and 3D preview fit | `renderFamilyTypeTable` / `renderCompositionDetail`; detached `BuildDetachedSelectionDetailHtml`; preview `fitPreviewImage`; shell CSS `.family-type-table` / `.preview-image-wrap` | `패밀리 구성` should show family types as a clear table, long type names should have enough dropdown width, and the small 3D preview should show the full image fitted inside the preview region while `크게 보기` keeps working. | User reported family types visually appeared as loose text (`# 타입 이름 1 ...`), the type dropdown was cramped and awkward, long type names could be clipped, and the small 3D preview showed a blank/white region while large view worked. Code had table markup but weak/fragmented styling, a narrow native select, and preview centering that depended on `inset:0` plus percent/transform positioning that can be unreliable in Revit's IE WebBrowser. | Backed up sources, added an overriding `family-type-panel`/`family-type-table` renderer and CSS in all 3 host projects, widened the type-parameter select to 420px+ flex layout, added IE-compatible preview bounds, changed small preview fit to pixel-calculated left/top centering, and added static guards. Static/contract checks, build, stage verification, WebBrowser harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Families / System Types | Second-row inline status filters beside `선택 항목 로드` / `선택 초기화` | `AppendFamilyInlineStatusFilterBar` / `AppendSystemInlineStatusFilterBar`; shell CSS `.inline-status-toggle`; shell JS `styleBrowserPane`; harness `CheckLayout` | The action row status filters should all remain visible when many states exist after a standard RVT scan, and the list/grid area should move below the wrapped toolbar instead of overlapping or clipping the filter buttons. | User reported the search-area filters were fixed, but the second action-row status filters still showed `...` when many filters appeared. Code confirmed `.inline-status-toggle` still had fixed max width, hidden overflow, and ellipsis. The runtime shell JS also reapplied the same clipping and fixed header height, so CSS-only changes could be undone inside Revit. | Backed up sources, changed all 3 hosts so action-row filters wrap with IE-compatible flex, removed inline ellipsis/hidden overflow, measured actual wrapped header height in shell JS, and set grid top from that measured value. Added dense audit fixture rows and static/harness guards for inline status clipping and action-row/grid overlap. Static/contract checks, build, stage verification, UI harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Families / System Types | Detail fingerprint difference summary and `상세 보기` table | `selectRow` -> `renderBasisMemo` -> `renderFingerprintDiffTable`; detached `BuildDetachedSelectionDetailHtml` -> `toggleFingerprintDiff` / `openFingerprintDiffModal`; harness `CheckBrowserDetailContent` | The detail basis/status area should show a short reason such as `타입 수 다름`, and the emphasized `상세 보기` button should open a table showing fingerprint difference rows such as type count and parameter value differences in both the main panel and detached detail window. | User reported fingerprint differences in the detail item were displayed as loose prose with little styling, `상세 보기` was non-responsive, and the basis/status text was too verbose. Code already emitted `data-diffrows` and had table-render helpers in the main page, but detached detail had `openDiffModal(){return false;}` and buttons depended on an in-page diff store id that was not available after copying HTML into the detached window. | Backed up sources, added concise diff labels/classification, changed diff buttons to carry encoded `data-diffraw`, emphasized the `상세 보기` button, added the diff table modal/parser/render path to detached detail HTML, removed the no-op diff route, added audit fixture diff rows, and extended static QA + WebBrowser harness to click the detail button and verify `fingerprint-diff-table` rows. Static/contract checks, build, stage verification, UI harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Admin Settings / Detached Detail | Standard source/list button layout, detail parameters, composite child table, 3D preview large view | `AppendAdminPane` / shell CSS; `OpenSelectionDetailWindow` -> `BuildDetachedSelectionDetailHtml`; JS `renderCompositionDetail`, `renderNestedChildTable`, `renderPreviewImage`; harness `CheckBrowserDetailContent` | Admin standard buttons should not overlap or waste space, detached detail should show the item name once with category, parameters should use the full detail width, complex families should list nested child category/family rows, and the 3D preview large-view action should open a modal. | User reported overlapped Admin Settings buttons, too much blank area, messy detached detail layout, duplicated family name, no 3D preview, non-responsive `크게보기`, and composite family child names not being shown in a structured table. | Backed up sources, changed Admin Settings CSS to a status-left/action-right grid with fixed-width action cards, kept detached detail title generic, preserved the inner family-name/category hero, forced parameter block full-width, added nested child category/family table rendering, made `크게보기` a clickable modal link, added composite-family audit fixture data, and extended static QA + WebBrowser harness to check nested table content, preview image markup, and modal opening. Static/contract checks, build, stage verification, UI harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Families / System Types | Auto-open detached detail on tab entry and detached detail layout | JS `setTab`/`window.onload` -> `kkyAutoOpenDetachedDetailForCurrentTab` / `kkyQueueAutoDetachedDetailOpen` -> `kkyfb:detail-window-open` -> `OpenSelectionDetailWindow` -> `BuildDetachedSelectionDetailHtml`; harness `CheckAutoDetachedDetailAction` | When the user enters Family Load or System Type Load and visible rows exist, the separate detail window should open automatically using the selected row or first visible row. The detail window should have stable, readable sections for identity/status, basis data, type/parameter data, and 3D preview. | Detail opening was still manual via the toolbar, so users could enter the list and think detail was broken or missing. The detached window was also narrow and the content cards/parameter/preview blocks felt cramped and uneven. | Backed up sources, added tab-entry/onload auto-open scheduling in all 3 host projects, kept normal row click as selection/sync, widened the detached window default/minimum size, and added detached-only CSS refinements for two-column detail content, parameter scrolling, and a fixed preview area. Added static regression guards and an `admin-system-with-data` harness scenario; static/contract checks, build, stage verification, UI harness for 32 scenario/year combinations, ProgramData install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Global Header | Project local/central file subtitle | `BuildDashboardHtml` -> `DisplayProjectSubtitleHtml` -> `TryResolveDisplayCentralPath` -> `TryGetCentralPath` | The subtitle under the browser title should stay compact: show project title, show local file name without a long inline path, expose the local path on hover, and when a workshared local file is open also show the central file name with the central path on hover. | Previous subtitle logic displayed the local project path inline next to the file name and did not render a separate central-file token in the header. Long paths could crowd the top status area and the user had to infer the central model from the local path. | Backed up sources, added HTML token rendering for local/central project files in all 3 host projects, resolved central paths from cached Revit worksharing data when available, added compact CSS for the subtitle tokens, added audit scenario local/central fixture paths, and extended static QA + WebBrowser harness checks for local/central visible names and tooltip path titles. Static/contract checks, build, stage verification, UI harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Automated UI Quality Gate | Static contract + WebBrowser click simulation | `Test-FamilyBrowserUiStatic.ps1`; `Test-FamilyBrowserUiContract.ps1`; `Invoke-FamilyBrowserUiAuditHarness.ps1`; `Invoke-FamilyBrowserQualityGate.ps1`; `FamilyBrowserDashboardHtmlForm.BuildDashboardHtmlForAudit` | Before build/install, all generated `kkyfb:` routes, browser-only JS buttons, major tab conditions, layout overlap checks, language/debug controls, missing-standard empty CTAs, and host action candidates should be checked without manual clicking. | Initial harness exposed a real 2019/2023 generated-JS regex break that made functions such as `openAdvancedFilter`, `setFilter`, and `loadFamilySelection` undefined. Harness also triggered `AnavRes.dll not found` popups by loading Revit UI native DLLs without Autodesk Shared dependency paths, and quality-gate report scripts had PowerShell array/parser issues. | Added audit contract, render seam, WinForms click harness, quality gate, native DLL popup suppression, Autodesk Shared dependency path discovery, hard process exit, PowerShell report fixes, and static guard for unescaped generated JS newline regex. Quality gate passed for 2019/2023/2025/2027. | Fixed |
| Admin Settings | `기준 RVT` / `표시 표준 목록` button layout | `AppendAdminPane`; `AppendStandardActionLink`; shell CSS `standard-action-layout` / `standard-action-row`; contract scenario `admin-standard-settings-layout` | The two right-side Admin Settings cards should read as workflow groups: primary setup first, secondary maintenance below, and trade/list utilities separated so users do not see a random inline button pile. | User screenshot showed the two cards rendering as mixed inline buttons with uneven wrapping. Primary actions, reset/delete, trade management, and list generation/export actions were visually mixed. | Backed up sources, replaced the mixed inline button pile with grouped standard-action rows, shortened labels, added shell CSS overrides so the external compact admin stylesheet does not collapse the layout, added static regression checks, and added a dedicated admin settings layout harness scenario. Static/contract checks, build, stage verification, and UI harness passed for 2019/2023/2025/2027. 2023/2025/2027 installed and verified; 2019 install is pending until Revit 2019 closes. | Fixed |
| Families / System Types | Status and discipline filter button layout after scan | `BuildDashboardHtml`; `.filterbar`; `.disciplinebar`; shell CSS `family-browser-shell.css`; harness `CheckLayout` | After standard RVT scan creates more status/trade filter buttons, the filter rows should wrap naturally and every visible label should remain readable instead of being clipped or ellipsized. | User screenshot and source inspection showed the filter rows were constrained to one line with `white-space: nowrap`, `overflow: hidden`, per-button `max-width`, and `text-overflow: ellipsis`. That made scanned-state filters vulnerable to text clipping when counts/status labels grew. | Backed up sources, changed inline and shell CSS in all 3 host projects to flex-wrap filter rows, removed max-width caps from filter buttons, changed label overflow to visible/clip, added static guards, and extended the WebBrowser harness to fail if rendered filters still use ellipsis or clipped visible text. Static/contract checks, build, stage verification, UI harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Families / System Types | Detail content and captured 3D preview | `OpenSelectionDetailWindow` -> `BuildDetachedSelectionDetailHtml`; JS `selectRow` / `requestInlinePreview`; harness `CheckBrowserDetailContent` | Selecting a standard family/system row should populate the detail surface with the item identity, status/category, family type/composition, captured parameter values, and any cached 3D preview captured during standard RVT registration. The detached detail window should not silently drop an existing preview PNG. | Precise registration did generate and store preview PNG paths, and rows emitted `data-previewpath`, but the detached detail window only relied on resolved file/URI preview source. If the embedded browser could not render the local PNG URI, the separate detail window showed the no-preview fallback even though the cached image existed. The automated UI harness previously checked clickability and layout, but not actual detail DOM contents. | Backed up sources, added detached detail inline data-URI fallback, added an audit PNG fixture and realistic family/type parameter summaries, and extended the WebBrowser harness to select an audit family row and assert detail name/category/type/parameters plus `preview-fit-image` or inline data URI. Static/contract checks, build, stage verification, harness, install, installed verification, and stage-vs-installed hash comparison passed for 2019/2023/2025/2027. | Fixed |
| Global | Dashboard action inventory | `BrowserNavigating` -> `DispatchOrRunDashboardAction` -> `RunDashboardAction` | All visible dashboard actions must route to a handler or known dynamic prefix. | Main dashboard actions route correctly. `manager-refresh` is handled inside the standard RVT manager sub-window, not the main router. | No code change needed. | OK |
| Global Browser Buttons | Nested WebBrowser custom-scheme buttons | Request draft `kkyfb-request:*`; standard family selection `kkyfb-select:*`; Standard RVT manager `kkyfb:*`; file guard `kkyfileguard://*`; sheet selector `kkysheet://*` | Buttons inside dashboard sub-windows must still execute if WinForms WebBrowser surfaces scripted custom-scheme navigation as `about:<scheme>...` instead of the direct scheme. | Main dashboard and detached detail already accepted `about:kkyfb:`, but nested request draft, selected-family update selection, Standard RVT manager, file guard, and sheet selector were more brittle. This could make buttons such as request add/remove/submit, update-selection apply, manager refresh/close, file-guard save/add/remove, or sheet select/cancel appear unresponsive in some WebBrowser navigation paths. | Backed up sources, added `about:` scheme parsing to all 3 host projects, added static QA guards, confirmed dashboard `kkyfb:` action extraction has 0 unknown actions, then passed static QA, build, stage verification, install, and installed verification for 2019/2023/2025/2027. | Fixed |
| Global Debug | Debug Log button / F12 | `debug-log` -> `ToggleDebugConsoleFromHost` -> `TryToggleDashboardDebugConsole`; JS `toggleDebug()` / `document.onkeydown` | Admin/internal-path users should see the dashboard debug overlay; users without that access should get the normal admin-only message or page-not-ready message, not a silent no-op. | The debug overlay was rendered from `showAdminUi && canViewAdmin`, while the host action used `CanSeeInternalPaths()`. Also, dashboard JS caught F12 even if `#fbDebug` was not present, so F12 could appear to do nothing. Host-side script invocation treated any `toggleDebug()` call as success even if no visible panel existed. | Backed up sources, aligned debug rendering to `CanSeeInternalPaths()`, added host-side panel-existence verification, changed JS so F12 toggles an existing panel or routes missing-panel F12 to `kkyfb:debug-log`, added static QA checks, then passed static QA, build, stage verification, install, installed verification, and stage-vs-installed SHA256 comparison for 2019/2023/2025/2027. | Fixed |
| Standard Target | Add trade target | `AddStandardDisciplineTarget` | New trade should become the active browser target so the next standard RVT registration applies to the new trade. | Policy active discipline changed, but `_browseDisciplineKey` was not synced, so registration could still target the previous browsed trade. | Synced `_browseDisciplineKey`, cleared all-slot cache, reset dashboard UI state in all 3 dashboard files. | Fixed |
| Standard Target | Switch standard mode | `SetStandardMode` | Integrated/separated mode switch should reset browser target to a valid slot. | Mode changed without normalizing the browse slot, leaving stale target risk. | Reloaded policy, resolved valid browse slot, reset cache/UI state in all 3 dashboard files. | Fixed |
| Family Load | Checked family selection | `ParseLoadableFamilySelections` -> `ApplyStandardFamilies` | Checked selections from different trades must remain distinct until multi-trade guard can warn/block. | Dedup key used only category/family, so same family name from another trade could be dropped before guard logic. | Added discipline-aware parse key in all 3 dashboard files. | Fixed |
| System Types | Checked system type selection | `ParseSystemTypeSelections` -> `ApplyStandardSystemTypes` | Checked selections from different trades must remain distinct until multi-trade guard can warn/block. | Dedup key used only kind/category/type, so same type from another trade could be dropped before guard logic. | Added discipline-aware parse key in all 3 dashboard files. | Fixed |
| Permissions | Legacy current permission check table | `CheckPermissionDiagnostic` -> hidden compatibility route | This route used to show Excel row/OX/source diagnostics, but Permissions / Guard is now intentionally file-guard-only. | Earlier audit row was stale after the UI removal. The visible Permissions tab no longer renders `permission-check`, role-list, permission-Excel, project-rule, or O/X diagnostic table controls. | Keep the hidden route only as compatibility/internal code. Do not list `permission-check` as a visible representative button unless that UI is deliberately restored. | OK |
| Global UI QA | Language switch, result dialogs, and text safety | `lang-ko` / `lang-en` -> `SetLanguage`; `ShowDashboardMessage`; `CurrentModelCheckResultDialog`; `FamilyLoadResultDialog`; `Test-FamilyBrowserUiStatic.ps1` | Korean/English switching, result popups, and generated HTML should not require the user to discover broken text, clipped result output, or stale detached-detail language by manual clicking. | User reported repeated misses around result windows, text breakage, and language switching. Static scan found that result dialogs used tight fixed sizing and the detached-detail refresh was called immediately after `DocumentText` replacement, before the new dashboard DOM was reliably ready. The audit MD also had stale representative actions. | Backed up sources, widened result dialogs, enabled explicit word wrap/shortcut-safe text boxes, made current-model result buttons width-aware, widened family-load result dialog/list space, and deferred detached-detail refresh until `BrowserDocumentCompleted`. Added static QA script for language routes, result wrapping, UTF-8 meta presence, and known regression tokens. | Fixed |
| Global UI QA | Runtime Korean -> English display-state localization | `lang-en` -> `SetLanguage` -> `RefreshLanguageSensitiveDisplayState`; `TranslateDashboardStandardText`; `TranslateDashboardPermissionText`; `TranslateDashboardTrackingText`; `TranslateProjectDisplayPlaceholder`; `DisplayDisciplineLabel`; harness `CheckLanguagePurity` | Switching an already-open Korean dashboard to English must immediately translate cached header status, unsaved-project text, window title, trade labels, list rows, and detail text without requiring refresh/reopen. | The prior harness rendered English from the start, so it missed stale strings created while Korean was active. Runtime switching translated only a limited `TranslateNote` subset: `설비`, `등록/미등록`, permission/tracking pills, `저장되지 않은 프로젝트`, and the native title could remain Korean. Stored row discipline labels and some audit row notes also retained Korean. | Added dedicated bidirectional status/project translators, immediate native title refresh, semantic discipline-key display localization, and exact row-note conversion. Audit scenarios now initialize English runs in Korean, the missing-RVT fixture uses the Korean unsaved-project placeholder, and the harness rejects visible Hangul plus Korean window titles. Focused 2025 passed 28/28; full five-target gate passed 112 `OK`, one 2021 runtime skip, zero failures, and 56 English transition results with zero language failures. ProgramData and installer were updated. | Fixed |
| Global UI QA | Dashboard / selection / file-guard message dialogs | `ShowDashboardMessage`; `ShowDashboardChoiceMessage`; `StandardFamilySelectionHtmlForm`; `FileGuardHtmlConfigurationForm`; `FamilyBrowserModernMessageDialog`; `Test-FamilyBrowserUiStatic.ps1` | Dashboard UI messages, standard-family selection validation, and file-specific guard confirmations should use the same modern, wrapped, Korean/English-aware dialog surface instead of scattered raw Windows message boxes. | Direct `MessageBox.Show` remained in the standard-family selection dialog and file-guard clear-all confirmation. `ShowDashboardMessage` also had raw MessageBox fallback branches and still instantiated the old nested `DashboardMessageDialog`. Search also found Revit-native `TaskDialog.Show` calls in external commands and native command guards; those are outside this dashboard UI pass. | Backed up sources, added shared `FamilyBrowserModernMessageDialog` to all 3 host projects, routed dashboard messages/choices, standard-family validation, and file-guard clear-all confirmation through it, and extended UI static QA to reject the previous raw routes. Static QA, build, stage verification, install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Current Model Check Export | Review Excel workbook difference columns | `CurrentModelCheckResultDialog.ExportRequested` -> `ExportCurrentModelCheckReviewList` -> `ProjectComparisonReviewExcelExportService.SaveReviewList` | Precise Current Model Check result dialog should keep an `Excel 추출` button that writes a formatted `.xlsx` workbook. The workbook must include difference item, standard value, project value, and difference summary for loadable families and system types. It should not expose review-result `.csv` export. | Loadable-family rows already populated difference columns from `FingerprintDifferenceDetails`, but system-type rows wrote blank difference columns. A follow-up mistakenly added `.csv` review-result export because "CSV" was misunderstood as export CSV rather than Revit family lookup CSV. User clarified that verification results do not need CSV export. | Kept the useful `.xlsx` difference-column fix, removed the Current Model Check `.csv` save-dialog option and CSV writer, normalized non-`.xlsx` output paths to `.xlsx`, updated the result message to Excel workbook export, and changed static QA so review-result CSV support fails the gate. Built/stage-verified/installed/installed-verified 2019/2023/2025/2027. | Fixed |
| Precise Family Fingerprint | Imported family CSV / lookup size table difference | `LoadableFamilyContentSignatureService.BuildCoreResultFromOpenFamilyDocument` -> `BuildLookupTableSignature` -> `FamilySizeTableManager.GetFamilySizeTableManager` -> `ProjectStandardComparisonService.ClassifySignatureDifference` | If a Revit family contains an imported lookup CSV / size table and the table names, column definitions, row values, or size definitions differ, the precise scan should treat the standard and project families as different. | User clarified that "CSV" meant the Revit family lookup CSV used inside `.rfa` files, not the Current Model Check review export CSV. Code already added `lookup-tables=` to the signature and classified `lookup tables` differences, but `ResolveFamilySizeTableManager` only searched for a one-argument manager API. Installed Revit API XML shows the actual supported API as `GetFamilySizeTableManager(Document, ElementId)`, so lookup table capture could silently return empty and miss differences. | Backed up sources, added `ResolveFamilySizeTableOwnerFamilyId(familyDocument)`, added the two-argument `GetFamilySizeTableManager(Document, ownerFamilyId)` reflection call in all 3 host projects, and added static QA checks for lookup signature inclusion, two-argument manager call, table name/body capture, `AsValueString` cell capture, and `lookup tables` comparison classification. Static QA passed, then 2019/2023/2025/2027 built, stage-verified, installed, and installed-verified. | Fixed |
| Families / System Types Detail | Lookup CSV presence and row/column difference display | `BuildLoadableDetailParameterPreviewText` / `BuildParameterPreviewText(StandardLoadableFamilySnapshotItem)` -> `BuildLookupCsvPreviewText`; `ProjectStandardComparisonService.BuildContentSignatureDifferenceDetails` -> `BuildLookupCsvDifferenceDetails`; detached detail `parseParameterPreview` / fingerprint diff modal | Detail windows should show whether a selected family has internal lookup CSV / size tables, how many rows/columns each table has, and fingerprint difference tables should make lookup CSV count/content mismatches readable instead of hiding them inside raw signature text. | The fingerprint path already captured lookup table content, but the detached detail surface did not summarize CSV presence or row/column counts. Difference details could show generic signature text instead of a clear `CSV 테이블` row, so users could not easily tell that the mismatch came from size-table CSV data. | Backed up sources, added lookup CSV summary parsing from content signature debug files and parameter preview text, rendered CSV rows in the detached parameter table, added explicit lookup CSV difference detail parsing for project-only/standard-only/count/content mismatch, translated Korean/English labels, added audit fixture `AUDIT_SIZE_TABLE`, and extended static QA + WebBrowser harness validation. Static/contract checks, Release build, stage verification, harness, ProgramData install, and installed verification passed for 2019/2023/2025/2027. | Fixed |
| Home | Dashboard overview buttons/cards | `AppendHomePane` -> `DashboardTabHref` -> `ShowDashboardTabFromAction` | Home should summarize registered standards, requests, processing state, recent updates, and missing setup. | Traced metric cards, trade board, recent activity, and next action links. Counts are read from `_loadableRows`, `_systemRows`, `_requestRecords`, `_unregisteredFamilyRows`, `_unregisteredSystemRows`, readiness report, and saved comparison/preflight/request paths. Home links only switch tabs; they do not mutate policy or project data. Request summary correctly shows `not loaded` when request store was not read yet. | No code change needed. | OK |
| Families | Search/filter/tree/detail/load actions | `AppendFamilyPane` -> JS `filterRows`/`setTreeFilter`/`prepareDetailWindowOpen`/`loadCheckedFamilies` -> `sync-family/*` / `sync-families-selected/*` -> `ApplySelectedStandardFamilyFromAction` / `ApplySelectedStandardFamiliesFromAction` -> `ApplyStandardFamilies` | Family list should be per selected trade, filters should remain client-side, detail should open in a separate window, and load actions should target the selected standard slot. | Traced search, status filter, tree filter, detail window, single load, and checked load. Search/filter/tree are browser-only and do not mutate data. Detail opens through `OpenSelectionDetailWindow`. Single/checked load includes `data-discipline-key`, blocks multi-trade checked selections, blocks category mismatch, resolves registration with selected discipline, and refreshes document shell after execution. | No code change needed for load/detail/filter path. | OK |
| Families / System Types | Detail window open action | `prepareDetailWindowOpen` / row `selectRow(..., true)` -> `kkyfb:detail-window-open` -> `OpenSelectionDetailWindow` | The toolbar `상세 항목` and selected row detail behavior should open a separate movable window from both Family and System Type tabs. | Current JS depended on the row-selection sync path and could issue preview/navigation side effects before the detail-window action reached the C# router. The main browser router also only accepted `kkyfb:` even though WebBrowser can surface scripted navigation as `about:kkyfb:`. Result: clicking `상세 항목` could update the inline detail panel but never open the detached window. | Backed up sources, added main dashboard support for `about:kkyfb:` action URLs, and overrode the detail-window JS so the open action is scheduled before selection preview work can swallow it. Built, stage-verified, installed, and installed-verified 2019/2023/2025/2027. | Fixed |
| Families / System Types | List command buttons and filters | JS `loadFamilySelection` / `clearFamilySelection` / `loadSelectedSystemType` / `clearSystemSelection` / `setFilter` / `setTreeFilter` / `setSystemTreeFilter` / row `selectRow` | Selection load/apply, selection reset, status filters, and tree/category filters should respond without being affected by detached detail navigation. | User reported `선택 항목 로드`, `초기화`, and filters gave no visible response. Traced generated onclicks: handlers existed, but the detached-detail patch wrapped `window.selectRow` and generated data rows still called `selectRow(this,true)`, so normal row selection could queue a detail-window navigation and leave following client-side button clicks in an unstable browser-navigation state. Load/apply permission false paths also returned silently. | Backed up sources, removed the `window.selectRow` wrapper, changed generated Family/System/Audit rows to `selectRow(this)` so row clicks only select, kept detached detail opening on the toolbar `상세 항목` button, and added permission-block alerts instead of silent returns. Confirmed old `selectRow(this,true)` / `kkyOriginalSelectRow` tokens are absent and handler functions remain present. Built, stage-verified, installed, and installed-verified 2019/2023/2025/2027. | Fixed |
| Families | Modeler update-visible rows | `FamilyRowsForCurrentMode` / `IsModelerVisibleFamilyRow` / `ModelerFamilyActionLabel` / `AppendFamilyModelerRow` / JS `loadCheckedFamilies` / `ApplyStandardFamilies` | If modeler view shows `Update Available`, the visible action and executable button should agree with intended behavior. | Code showed update/tracking rows to modelers and labeled them `Update from standard` / `Refresh tracking`, but JS load actions only allow `LoadAvailable`. The selected-load plan also treats already loaded families as skipped and says existing updates are handled from Admin Settings only, so the modeler UI implied an action that could not execute. | Kept governance behavior: modelers can still see update/tracking rows for awareness, but only `LoadAvailable` family rows get an enabled checkbox. Update/tracking detail action text now says admin update/tracking refresh is required. Built, stage-verified, installed, and installed-verified 2019/2023/2025/2027. | Fixed |
| System Types | Search/filter/tree/detail/apply actions | `AppendSystemPane` -> JS `filterRows`/`setSystemTreeFilter`/`prepareDetailWindowOpen`/`loadSelectedSystemType` -> `sync-system/*` / `sync-systems-selected/*` -> `ApplySelectedStandardSystemTypeFromAction` / `ApplySelectedStandardSystemTypesFromAction` -> `ApplyStandardSystemTypes` | System type list should be per selected trade, detail should open in a separate window, selected apply should validate dependencies/duplicates/category/selected standard target without full-catalog re-scan, and confirmation/result windows should use the browser HTML operation style. | Re-traced search, status filter, tree filter, detail window, single apply, checked apply, selected preflight, execution report save, post-apply check, and refresh path. Search/filter/tree remain browser-only. Checked selections with multiple trades are blocked. Category mismatch and duplicate risk are blocked in JS and rechecked in C#. Found regression: multi-selected system-type apply used full `BuildReport(...)` before filtering selected rows, and post-apply verification repeated the same broad pass. Result/confirmation UI was still plain WinForms/text in several family/system apply paths. | Backed up sources, changed selected system-type apply to `BuildSelectedSystemTypePreflightReport(...)`, merging per-selected `BuildSelectedReport(...)` outputs and mapping progress as selected-row progress instead of full scan progress. The same selected-only route is used after apply. Family load and system type apply confirmation/result windows now use `FamilyBrowserOperationHtmlDialog` with HTML cards/tables/report actions. Post-result shell/dashboard refresh now shows progress. Static QA rejects the legacy multi-selected full `BuildReport(...)` branch and legacy family result dialog. Full quality gate passed 2019/2023/2025/2027 with 72 harness scenarios and 0 failures. | Fixed |
| System Types | Bulk apply from Audit tab | `AppendFingerprintDifferenceSection` -> `ToolActionLink("apply-systems")` -> `RunDashboardAction` -> `ApplyStandardSystemTypes(null, null, null, _browseDisciplineKey)` | Admin bulk apply should only run for current selected standard target, require governance permissions, revalidate the preflight, and avoid applying duplicate/category-risk rows blindly. | Traced visible `Bulk Apply Standard System Types` action. The link is rendered only as enabled with `ManagePolicy` and `ApplySystemTypes`; handler rechecks both permissions plus readiness, loads the selected standard registration, rebuilds system preflight, filters by approved standard list for the selected browse discipline, blocks duplicate risk/category/dependency blockers, confirms counts, executes supported create/overwrite/consolidate actions, saves preflight/apply reports, stamps tracking only on clean mutation, and refreshes dashboard. | No code change needed. | OK |
| Requests | Request create/status/delete/attachments | `AppendRequestPane` -> `RunDashboardAction` -> `CreateRequestDraft` / `RefreshRequestsAndAudit` / `UpdateRequestStatusFromAction` / `DeleteRequestFromAction` / `OpenRequestAttachmentFolderFromAction` -> `FamilyBrowserRequestStore` | Request actions should respect store availability, permissions, creator/admin rules, and attachments. | Traced register, refresh, family/system/update request cards, status transitions, delete, attachment-folder open, and request store readiness. Creation requires `CreateRequest`, verifies writable shared request store, prompts only admins to repair storage, blocks local C/AppData/UserProfile fallback, copies attachments into a request-specific folder, writes JSON/mail draft/manifest, and reloads requests. Status transitions require `SubmitRequest` for reopen/submit and `ApproveRequest` for review/approve/reject/complete. Delete is visible/executable only for creator or admin and store deletion is constrained to files/folders inside the request-store root. Attachments can be opened only by admin/internal-path users. Refresh loads the store only on Home/Requests or forced refresh, so startup does not probe the request path. | No code change needed. | OK |
| Requests | Request composer sub-window | `RequestComposerForm` -> `kkyfb-request:add-files` / `remove-files/*` / `submit` / `cancel` -> `CreateRequestDraft` | Request popup buttons should collect request text and attachments without exposing request-store paths, then return validated values to the store writer. | Traced composer actions. `add-files` captures current form state and appends unique attachment files, `remove-files/*` validates indexes before removing, `submit` requires title/body, and `cancel` returns without writing. Discipline chips are client-side only. On OK, `CreateRequestDraft` rebuilds the `FamilyBrowserRequestRecord`, copies attachments into the request-specific folder, writes JSON/mail draft/manifest, and clears pending request attachments. | No code change needed. | OK |
| Audit / Model Check | Current model check / tracking / stamp | `AppendCheckMaintenanceWorkspace` -> `refresh-target/*` / `preflight-target/*` / `sync-families-target/*` / `sync-systems-target/*` / `stamp` / `tracking` -> target handlers / `RefreshDashboard` / tracking services | Model checks should use selected standard target and avoid stale cross-trade cache. | Targeted actions decode the selected slot token, resolve the exact policy slot, set `_browseDisciplineKey`, and then run current model check, system diagnostics, existing-family update, or existing-system update against that selected target. Found one tracking-state bug: `RefreshDashboard` loaded tracking before `_registration` was set, so a saved tracking catalog from another standard source could miss the `source mismatch` warning. Found another related state bug: after Current Model Check cleared the protected-deletion dirty marker, tracking text could remain `Current Model Check required` until a later refresh. | Added `UpdateTrackingState(trackingCatalog)` after `_registration` is assigned and after `ClearCurrentModelCheckRequired(doc)` in all 3 dashboard files. Built, stage-verified, installed, and installed-verified 2019/2023/2025/2027. | Fixed |
| Audit / Model Check | Existing family update selection dialog | `ApplyExistingStandardFamiliesForTarget` -> `ApplyStandardFamilies` -> `PromptExistingFamilyUpdateSelection` -> `StandardFamilySelectionHtmlForm` / `kkyfb-select:apply,cancel` | Admin-targeted existing-family update should let the user choose exact plan rows without merging same-named families from different categories. | Dialog displayed category and family, but selection payload used only family name. If the same family name existed in more than one category, selecting one row could include all executable plan rows with that family name. | Backed up sources, changed selection payload and filtering to use `BuildFamilySelectionKey(category, family)` in all 3 dashboard files. Built, stage-verified, installed, and installed-verified 2019/2023/2025/2027. | Fixed |
| Unregistered Families | Detection/export/request actions | `AppendLoadableRowsFromComparison` / `RebuildUnregisteredRowsFromProjectLightInfo` -> `AppendUnregisteredFamilyPane` -> `unregistered-families-export` / `request-unregistered-family/*` -> `ExportUnregisteredFamilies` / `CreateRequestFromUnregisteredFamilyAction` | Unregistered families should be generated from current model check, available for admin review/export, and convertible into a request without retyping row metadata. | Traced both generation paths. Current Model Check routes `ProjectOnly` loadable families into `_unregisteredFamilyRows` and excludes them from the normal family load list. Lightweight/startup classification also compares current project family names against approved standard Excel/JSON list entries and hides nested helper families. UI showed count/table and exported to Excel with `ManagePolicy`, but there was no row-level request action; users had to retype discipline/title/reason manually in the request composer. | Added row-level `요청` action gated by `CreateRequest`, encoded row metadata into `request-unregistered-family/*`, added UI-only router prefix, and prefilled `CreateRequestDraft` with discipline, title, and reason. Built, stage-verified, installed, and installed-verified 2019/2023/2025/2027. | Fixed |
| Unregistered Systems | Detection/export/request actions | `AppendSystemRowsFromComparison` / `RebuildUnregisteredRowsFromProjectLightInfo` -> `AppendUnregisteredSystemPane` -> `unregistered-systems-export` / `request-unregistered-system/*` -> `ExportUnregisteredSystemTypes` / `CreateRequestFromUnregisteredSystemAction` | Unregistered system types should be generated from current model check, available for admin review/export, and convertible into a request without retyping row metadata. | Traced both generation paths. Current Model Check routes `ProjectOnly` system types into `_unregisteredSystemRows` and excludes them from the normal system type load list. Lightweight/startup classification compares current project system type class/name/category against approved standard Excel/JSON list entries. UI showed count/table and exported to Excel with `ManagePolicy`, but there was no row-level request action; users had to retype system class/type/category manually in the request composer. | Added row-level `요청` action gated by `CreateRequest`, encoded row metadata into `request-unregistered-system/*`, added UI-only router prefix, and prefilled `CreateRequestDraft` with discipline, title, and reason. Built, stage-verified, installed, and installed-verified 2019/2023/2025/2027. | Fixed |
| Admin / Standard Management | Standard RVT manager and standard list actions | `AppendAdminPane` / `StandardRvtManagerHtmlForm` -> `RunDashboardAction` -> standard RVT/list handlers | All standard operations must be scoped to the selected trade/integrated target. | Traced Standard RVT manager sub-window actions: `manager-refresh` stays inside the popup, `standard-rvt-manager-close` closes it, `browse-discipline-*` switches the selected target and redraws, maintenance buttons close the popup and dispatch through the main dashboard router. Traced standard list and trade actions: add/rename/delete trade, connect Excel, save template, export from RVT, save fingerprint audit, clear list, and reset selected standard RVT. These handlers resolve `ResolveBrowseSlot(policy)` / selected slot, enforce `RegisterStandard` or `ManagePolicy`, update slot-scoped policy paths, clear dashboard/modeler caches, and save the standard list index where needed. | No new code change needed for the traced Standard RVT manager / standard list route. Existing-family update selection bug was recorded separately because it belongs to the target update dialog. | OK |
| Admin / Standard RVT Precise Scan | Full/selected precise scan dialog guard and diagnostics | `standard-rvt-full-precise` / `standard-rvt-selected-precise` -> `RefreshRegisteredStandardRvt` / `RefreshSelectedStandardFamiliesPrecise` -> `StandardLibraryRegistrationService` -> `FamilyThumbnailConstraintDialogGuard`; Current Model Check deep capture | During precise scans and Current Model Check family editing, non-destructive warning/error dialogs with an enabled OK button should be acknowledged and recorded. Dialogs that offer only destructive Delete Instance/Delete Type/Delete Constraints/Remove Constraints choices plus Cancel should be cancelled and recorded. The result should expose a reviewable on-demand XLSX error/action log. | The previous guard still depended on recognized warning text, so `Opening not cutting anything` and other unseen OK-only family-edit warnings could remain open. A native Delete button carrying the numeric `IDOK(1)` control ID was also a safety edge because the control ID could be mistaken for semantic OK. | Active family-edit scopes now use actual enabled button topology: visible OK/Confirm/Continue wins for any warning text; no-OK delete/remove-only choices plus Cancel select Cancel. `Opening not cutting anything` has an explicit reason. Destructive button text overrides a misleading IDOK control ID. Standard-scan summaries identify affected families, and Current Model Check exports a localized `스캔경고` / `ScanDialogs` sheet only when the user clicks Excel export. Five-target build/Stage, full 136-scenario UI gate, and post-safety compiled-host audits passed; actual native Revit dialogs remain Next Work Queue item 25. | Needs Revit Check |
| Admin / Standard List | Excel sheet selection sub-window | `ConfigureStandardListExcel` -> `SelectStandardListSheetName` -> `StandardListSheetSelectionHtmlForm` -> `kkysheet://select/{index}` / `cancel` | When connecting an Excel standard list, the user should choose a worksheet safely and the selected sheet should be passed to materialization. | Traced sheet popup. It loads workbook sheet names, defaults to the current sheet or first sheet, validates the selected index before setting `SelectedSheetName`, returns cancel without policy write, and falls back to the older text prompt only when sheet enumeration fails or returns no sheets. | No code change needed. | OK |
| Families / System Types | Standard setup empty-state classification | `AppendFamilyPane` / `AppendSystemPane` -> `IsActiveStandardListRegistrationRequired` -> `AppendStandardListRegistrationRequiredRow` / `AppendStandardRegistrationRequiredRow` -> setup action | The selected target must distinguish three states: RVT missing, RVT connected but scan needed, and RVT ready but approved standard list missing. When only the list is missing, Admin Mode must show the list message/action and must not show the RVT registration message/action. | The list-specific renderer and route already existed, but `IsActiveStandardUnavailable(_registration)` ran first. If the dashboard registration record was temporarily unavailable while the selected policy slot still contained RVT connection metadata, the UI incorrectly fell back to `표준 RVT를 먼저 등록해주세요`. The audit contract also lacked a Family-tab missing-list scenario and its missing-RVT fixture incorrectly set scan-needed. | Added selected-slot-aware `IsActiveStandardListRegistrationRequired()`, moved the list branch before the registration-record fallback in both Family and System panes, changed the prompt to `등록된 표준 RVT의 표준 목록을 연결해주세요`, and kept a distinct generic empty state when both setup artifacts exist but no rows are ready. Added Family/System Korean/English CTA/message isolation checks, corrected the audit fixture, and stabilized queued-search timing. Static/contract checks passed; five-target build/stage and install verification passed; all new setup scenarios passed across installed runtimes and the stable 2019 rerun passed 20/20. | Fixed |
| Admin / Hidden Governance Routes | Legacy/internal admin actions with no current visible button | `RunDashboardAction` cases for `request-store-*`, `deployment-*`, `security-*`, `permission-excel-*`, `project-rule-*`, `settings-reset-all`, `policy-*`, `temporary-register`, `open-managed-root`, `open-requests`, `output` | 1.0 should keep Standard Admin Settings focused on standard RVT/list setup, while operational governance controls live only where the user wants them. Removed manual-managed-folder flows should remain compatibility-only. | Re-scanned literal generated `kkyfb:` hrefs from the representative 2019-2023 dashboard file. Real operational handlers existed but had no visible tab-scoped entry point. User corrected that Admin Settings was intentionally split by tabs, then later clarified that Permissions / Guard should not show role lists or permission Excel. Explicit removed-flow handlers still only show removed-flow messages. | Request controls remain on Requests; deployment/path/reset controls remain on Help / Technical Details. Permissions / Guard now renders only file-specific guard status/config/export. Deprecated `security-*`, `permission-excel*`, `project-rule*`, `policy-*`, `temporary-register`, `team-shared-test`, and `request-store-local` remain hidden compatibility routes with no visible Permissions tab entry point. Built, stage-verified, installed, and installed-verified 2019/2023/2025/2027. | Fixed |
| Permissions / File Guard | Configure/export file-specific guard | `AppendPermissionPane` -> `file-guard-config` / `file-guard-export` -> `ConfigureFileGuardPolicy` / `ExportFileGuardPolicy` -> `FileGuardHtmlConfigurationForm` | Permissions / Guard tab should be dedicated to file-specific native command guard settings only. | User requested removing everything in the role-list / permission-Excel area and using this tab only for file-specific permission application. The tab previously rendered current user role, admin rule, role lists, permission Excel controls, project rules, current permission check, and an O/X diagnostic table. | Backed up sources, removed the role/Excel/project-rule/admin-rule/current-permission-check/O-X table render path from `AppendPermissionPane`, kept only current model, file-specific guard status, `파일별 권한 적용 설정`, and guard status export. Built, stage-verified, installed, and installed-verified 2019/2023/2025/2027. | Fixed |
| Permissions / File Guard Runtime | Save policy, Admin OFF, Load Family/type-change interception | `ConfigureFileGuardPolicy` -> `FamilyBrowserStandardPolicyStore.SetFileGuardPolicy` -> `FamilyBrowserNativeCommandGuardService`; Revit `FamilyLoadingIntoDocument`; protected-content updater | A saved target RVT must block Revit Load Family and family/type changes when an authorized administrator switches Admin Mode OFF. Local workshared copies and mapped-drive/UNC aliases must match the guarded central RVT. | Current PC could not reach either homepage-managed path, so policy save threw and the HTML error hid the real exception behind `(managed log unavailable)`. On another PC, Admin OFF synchronously wrote user settings to the network share, rebuilt the full browser shell, recursively traversed the Autodesk ribbon, and fast guard contexts omitted the workshared central path. Path comparison also used only `Path.GetFullPath`, so mapped-drive, hostname UNC, and IP UNC spellings were distinct. These conditions could make the toggle slow and let a differently named local copy miss the guard target. | Backups `_backups\file-guard-runtime-enforcement-20260715-111457` and `_backups\file-guard-path-alias-20260715-115806`. Moved user language/Admin settings to `%LOCALAPPDATA%`, added managed-folder preflight and local diagnostic fallback, handed saved policy directly to the runtime guard, removed full shell reload/ribbon traversal, and cached workshared central identity. Mapped drives now expand through cached `WNetGetConnection`; hostname/IP UNC aliases also compare by exact share-relative path while different shares remain isolated. Static/contract checks, five-target Release build/Stage verification, 2,000-row performance gate, prior `136 OK` UI scenarios, and the new focused alias/runtime-policy harness passed. Actual Revit native-command cancellation on the connected PC remains queue item 23. | Needs Revit Check |
| Permissions / File Guard Runtime | Nested-only family standalone placement | `FileGuardHtmlConfigurationForm` -> `BlockNestedOnlyStandalonePlacement` -> precise standard snapshot schema 6 -> `FamilyBrowserNestedOnlyPlacementCatalogStore` -> `FamilyBrowserNativeCommandGuardService.CollectNestedOnlyStandalonePlacements` | Per registered RVT, Admin OFF should prevent direct project placement of a family used only as a nested child in the selected standard. Placing its parent family must remain valid, and disabling the option or using Admin ON must allow direct placement. | Existing snapshots did not record whether a nested child also had project-level standalone instances, so nested metadata alone could not distinguish nested-only from dual-use families. Blocking by family name alone could also collide across categories. An absent catalog could also remain negatively cached for up to 30 seconds if the policy was enabled before a new scan. | Added standalone usage capture, precise-scan completeness gate, category+family catalog, atomic managed sidecar/local cache, explicit File Guard checkbox/persistence/export, addition-only `FamilyInstance` updater check, `SuperComponent` exemption, Admin bypass, rollback/audit metadata, static guards, semantic harness tests, and immediate catalog-cache invalidation after full/selected standard scan attachment. Legacy/incomplete scans fail open. `2019/2021/2023/2025/2027` Release build/Stage, static/contract, 2,000-row performance/cache, and full IE harness passed with zero failures (`136 OK + 1 expected 2021 runtime SKIP`). Dedicated policy: `FAMILY_BROWSER_NESTED_ONLY_PLACEMENT_POLICY.md`. Actual Revit rollback remains queue item 28. | Needs Revit Check |
| Debug | Debug log and status actions | `fbDebugTool` / `fbDebugFab` / F12 -> `debug-log` -> `ToggleDebugConsoleFromHost` / `TryToggleDashboardDebugConsole` / `BuildRuntimeDebugInfo` / `WriteDashboardRuntimeDiagnostic` | Debug actions should expose useful paths/state only to internal-path/admin users and should never fail silently. | Earlier debug guard work hid the overlay from normal modelers, but the MD row was stale and still described the render gate as `showAdminUi && canViewAdmin`. The 2026-07-06 recheck found the remaining mismatch: host access used `CanSeeInternalPaths()` while JS could still catch F12 without a rendered panel. | Current source now uses `canUseDebug = CanSeeInternalPaths()`, renders debug buttons/payload only with that gate, verifies `#fbDebug` from the host before treating a toggle as successful, and routes missing-panel F12 to the host `debug-log` action so the user gets a message. Static QA, build, stage verification, install, installed verification, and stage-vs-installed hash comparison passed for 2019/2023/2025/2027. | Fixed |
| Global Managed Data / Homepage | Startup path resolution, standard snapshot read, project scan cache restore, V2 reference integrity | `DeferredInitialOpenRefresh` -> `PrepareStartupPreload` -> `FamilyBrowserDeploymentBootstrapService.TryApply` -> `FamilyBrowserStandardPolicyStore`; `InvalidatePreparedDashboardDataAfterManagedPolicyChange`; `StandardLibraryRegistryStore`; `ProjectSnapshotStore`; `FamilyBrowserDataLoader`; `Test-FamilyBrowserManagedData.ps1` | The first browser open must use the homepage-selected management root. Standard scans and per-project comparison data must be stored under that root and read back with intact references. Project aliases must not confuse unrelated same-name files. | Startup preload previously loaded policy before homepage bootstrap, then reused that stale prepared result after deferred or manual homepage path/profile/URL/security changes. Project alias save/read also returned early because `_workspaceRoot` is intentionally empty, despite the effective managed Projects folder being available through runtime policy. Live homepage candidates are currently unreachable on this PC, so real shared contents remain a runtime check. | Bootstrap now runs before preload policy/data reads; every homepage path/policy refresh route invalidates prepared caches through one helper. Project alias persistence resolves the effective managed Projects folder and requires a matching current file stamp. Added a read-only managed-data integrity audit with normal and broken-reference fixtures, optional quality-gate integration, static guards, five-target build/install verification, and full UI/performance regression. Re-run against a reachable homepage folder for the remaining live-data check. | Needs Revit Check |
| Global Performance / Cache | Startup, browser-index load, row/detail/image rendering, stale-cache mutation guard | `RenderStartupShellOnly` -> `PrepareStartupPreload` -> `IFamilyBrowserDataLoader`; V2 manifest/index/detail/thumbnail/project/row cache; `replaceDashboardPaneHtml`; 150-row window; `ValidateCurrentSourceRevision` | Open a responsive shell immediately, read compact validated local data without rescanning RVT, keep network source authoritative, search 1,000 rows quickly, load detail/3D only for selection, and never mutate a Revit model from stale cached standard data. | Previous startup synchronously read/merged broad snapshot data, repeatedly resolved image paths, rendered every row/detail into full HTML, and the performance gate incorrectly combined fresh-process host DLL load with user-visible shell time. | Added shared V2 loader/models, hash/revision manifest, atomic local cache, offline fallback, background generation cancellation, compact snapshot/list preparation, one-time thumbnail index, lazy detail/preview hydration, partial pane update, 150-row visible window, source-revision mutation guard, startup-shell audit seam, and 2,000-row performance gate. Final five-target quality gate, ProgramData verification, and installed hash comparison passed; Revit 2021 runtime remains unavailable on this PC. | Fixed |
| Global Performance / DOM Virtualization | True fixed-shell row injection | `BuildVirtualRowPayloadJson` -> embedded `family-browser-row-window.js` -> `window.KKYFB.setRows(...)` / `setRowsFromJson(...)` -> active 150-row DOM window | Keep all indexed rows searchable while creating only the active 150-row slice in the IE WebBrowser DOM. Paging, checked rows, selected detail, saved UI state, partial pane refresh, and Family/System load/apply payloads must continue to work across windows. | The previous 150-row mechanism only hid off-window rows after C# emitted every row as HTML, leaving 1,000 DOM rows and approximately 1.29-1.31MB audit HTML. Selection helpers also only knew about currently rendered checkboxes, so true virtualization required a full-data checked-row adapter. | Backed up all three hosts/tests, added compact 31-field row payloads with safe JSON escaping and stable keys, embedded a shared ES5 virtual row store, generated at most 150 real rows, preserved cross-page checks and load/apply adapters, restored saved selections by moving to the owning page, cached search text/window signatures, and fixed Clear Selection to clear both checks and selected state. Harness now rejects hide-only DOM, tests page 1 -> 2 -> 1 persistence, and records total/DOM/visible separately. Five-target quality gate, ProgramData install, installed verification, and installer build passed. | Fixed |
| Families / System Types | Resizable list columns and numbered 150-row paging | fixed header/body `colgroup` -> `family-browser-row-window.js` `ensureColumnResizers` / `goToRowWindowPage`; bottom `dashboardStatusBar` paging controls | Users must be able to drag every list-header boundary while header/body widths remain aligned. Lists over 150 rows must expose obvious Previous/Next and direct numbered-page navigation near the screen-transition status instead of a small temporary-looking search-line control. | The old pager was embedded beside the search input and exposed only arrow glyphs plus a row range. Fixed column widths could not be adjusted, and changing one table alone would risk header/body misalignment. Updating the outer status bar text would also destroy nested controls if paging were moved there unchanged. | Moved paging into a separate 50px bottom status region with independent `dashboardStatusText`, localized Previous/Next, current/total summary, direct page numbers, ellipsis, and row range. Added synchronized header/body drag resizing, minimum widths, per-tab/column-count persistence, and audit seams. The 2,000-row IE harness verifies direct page 2, active numeric state, cross-page checks, handle count, and exact header/body width equality. Five-target gate/install and installed hash comparison passed. | Fixed |
| Model Check / Admin Standard Target | Target switch scroll, active state, and standard action layout | `browse-discipline-*` / `discipline-*` -> `SetBrowseDiscipline` / `SetActiveDiscipline` -> `CaptureDashboardUiState` / `RefreshDocumentShellOnly`; `AppendDisciplineLink`; Admin Standard action markup/CSS | Changing the Model Check target must preserve the current scroll position. Admin Standard must show the actual current target as pressed, keep trade management next to the target selector, and use consistent action-button rows. | `SetBrowseDiscipline` explicitly cleared `_lastDashboardUiStateJson` and skipped restore even on the audit tab. Admin trade chips used `policy.ActiveDiscipline` while the displayed current target used `_browseDisciplineKey`, so the label and pressed chip could diverge. Trade management lived inside Baseline RVT, and final IE-safe CSS reverted action rows to auto-width inline buttons. | Audit target changes now capture/preserve UI state while Family/System target changes retain reset behavior. Admin separated mode resolves active state from the current browse target, trade management moved into `admin-trade-control`, and Baseline/List rows use equal flex widths/heights with full-width terminal actions. Static contract and WebBrowser harness enforce scroll round-trip, one active target, management placement, and row dimensions. The current five-target build was installed and installed-verified on 2026-07-10. | Fixed |
| Global UI QA | Structured result/error/confirmation message body | `ShowLoggedError` / `ShowDashboardMessage` / `ShowDashboardChoiceMessage` -> shared `FamilyBrowserModernMessageDialog` -> `FamilyBrowserMessageHtmlRenderer`; `BuildHtmlForAudit`; `RunMessageBodyAudit` | Common dashboard result/error/confirmation popups should present a clear headline and visually separated cause, next action, administrator, and technical information instead of a large undifferentiated text block. Long details must stay readable, Korean/English must remain clean, and closing a choice dialog must not approve the action. | The custom dialog shell was already WinForms, but its body was one readonly multiline `TextBox` populated from newline-delimited `FamilyBrowserFriendlyError.ToDialogMessage(...)`. This made every heading and paragraph visually equal, crowded long errors, offered no structured copy action, and custom `X`/`Esc` returned the default affirmative result for Yes/No dialogs. | Moved the dialog into the shared UI project, replaced the body with a compact IE-compatible HTML renderer, added severity hero plus section cards, metadata rows, internal technical scroll, UTF-8 declarations, and copy-details action. `X`/`Esc` now return No for choices. Static checks and eight Korean/English message harness runs enforce the structure, language, overflow, and no-body-clickable contract. Five-target quality gate, ProgramData install, installed verification, and stage/installed hash comparison passed. | Fixed |
| Global UI QA | Main dashboard default window size | constructor -> `ResolveStartupScreen()` -> `ApplyRecommendedWindowSize()` -> `GetRecommendedStartSize()` / `GetRecommendedMinimumSize()` | Family Browser should open at the same practical size as KKY Tool: `1400 x 900`, shrink to at most 93% of the active working area, retain `1100 x 720` minimum where possible, and remain centered on the Revit monitor. | The dashboard used its own much larger responsive targets: up to `1680 x 960` on ordinary wide screens and `1780 x 1040` on larger screens, with unrelated adaptive minimum floors. | Replaced all three host calculations with the KKY Tool rule while preserving the existing Revit-monitor resolver, explicit centering, normal window state, and user resize behavior. Static guards lock the width, height, minimum size, and centering tokens. Five-target quality gate/install passed; actual workstation DPI appearance remains a short Revit runtime confirmation. | Needs Revit Check |
| Global UI QA | Persisted-state immediate refresh after actions | state-changing handler -> policy/registration/list save -> `RefreshDocumentShellOnly` -> live store load -> `RenderDashboard` with UI-state restore | Every successful dashboard change must appear immediately without the user pressing Refresh. Startup performance cache may accelerate only the first render and must never overwrite later saved state. | The handlers already saved and called `RefreshDocumentShellOnly`, but that method preferred `_startupPreloadResult.Policy` and prepared slot data whenever the browser had been opened once. A trade rename therefore wrote the new display name correctly, then immediately rendered the old startup policy; manual Refresh used the live store and finally showed the change. The same stale-source risk affected mode, add/delete, RVT/list, security, guard, project-policy, and request-store mutations. | Added an initial-render-only preload permission passed through the complete shell refresh and reset in `finally`. All normal/post-action refreshes now load live policy and current slot metadata; prepared startup slots are gated out. Runtime diagnostics record `startup-preload` vs `live-store`. Static QA enforces this contract and immediate refresh presence across 18 persisted-state mutation methods. Full five-target gate/install passed; the reported rename sequence remains a final external-PC Revit confirmation. | Needs Revit Check |
| Families / System Types | Provisional load/apply state and save/sync finalization | load/apply result -> `RecordLoadableFamilyOperation` / `RecordSystemTypeOperation` -> pending operation queue -> dashboard overlay; `DocumentSaved` / `DocumentSavedAs` / `DocumentSynchronizedWithCentral` -> commit verification; `DocumentClosing` / `DocumentClosed` -> discard | A successful in-memory load/apply must stop appearing as `Load Available` immediately, but must not become a persisted completion until Revit successfully saves or synchronizes the document. Closing without saving must discard the provisional state; a cancelled close must retain it. | Operation entries already carried `PendingSaveOrSync`, but the dashboard did not consume that state and ordinary Save/Save As did not finalize it. The list could therefore keep showing `Load Available`, and operation completion did not distinguish an in-memory mutation from a persisted document change. | Added a per-open-document pending queue, source-scoped Family/System row overlay, disabled temporary rows, successful Save/Save As/Sync commit handlers, live-document presence verification, close-completion discard, post-commit cache invalidation/refresh, and precise-scan deferral while temporary state exists. Failed/cancelled save status does not commit. Static/contract, five-target build/stage, performance, and 136-result IE harness passed; real Revit event ordering remains the final runtime check. | Needs Revit Check |
| Families / System Types | Prepared trade target switching from list chips | `AppendDisciplineLink` -> JS `beginBrowseDisciplineSwitch` -> `kkyfb:browse-discipline-*` -> `SetBrowseDiscipline` -> prepared slot/project scan restore -> row-cache rebuild | Clicking a ready trade below the Family/System search field must switch both the pressed chip and the actual rows to that trade, with the same target semantics as selecting it in Standards and without an unnecessary network reload. | `SetBrowseDiscipline` changed `_browseDisciplineKey`, but normal shell refresh was forbidden from consuming already prepared non-startup slots. It reread source data synchronously, only the startup-selected trade had a preloaded project comparison, and the persistent row-cache key could be calculated before an alternate trade's saved project scan was restored. During the wait, the old trade's rows remained visible, making the chip look ineffective. | Added a separately scoped prepared-slot permission for explicit target switching while retaining live-policy reads, restored the selected trade's saved project scan before row-cache lookup, cached the restored scan/stamp, bumped row-cache generation to `v7-trade-switch-project-restore`, and invalidated prepared slots after RVT/list mutations. IE feedback immediately activates the clicked chip and hides prior-trade rows. Static guards and the real IE harness switch inactive trade chips in Family/System scenarios and verify one active target plus no old-target rows. Five-target build/Stage and the full `136 OK` gate passed; a two-real-trade Revit data check remains queue item 26. | Needs Revit Check |
| System Types / Detached Detail | Revit-style compound layer table and persisted display unit | `CaptureSystemTypeLayers` -> `StandardSystemTypeLayerSnapshotItem` -> `SystemTypeDetailSummaryService @layer` -> `FamilyBrowserSystemRoutingUnitUi.renderSystemRoutingLayers`; `kkyfb:measurement-unit/*` -> `SetMeasurementDisplayUnitPreference` -> `FamilyBrowserMeasurementUnitPreferenceService` | Compound layers should resemble Revit Edit Assembly: exterior/interior direction, function, material, thickness, core boundaries, structural/variable metadata. Routing criteria and layer thickness should both support synchronized `mm` / `in`, default to `mm` on first use, and restore the last unit after closing/reopening. | Layer details were flattened into `Function / Material / Thickness` strings, so the composition was difficult to scan and core/structural/variable semantics were lost. The routing selector always initialized to `mm`, did not persist, and did not control layer thickness. Final visual QA also found that the audit-only detached canvas forced hidden blocks visible, making a legacy empty preview card and Family composition block appear in the captured System detail. | Added structured `@layer` records with raw feet plus core/structural/variable flags, a Revit-style layer table with legacy-row fallback, synchronized routing/layer unit controls, LocalAppData atomic preference persistence, and main/detached detail synchronization without rebuilding the document. Removed the audit CSS visibility override, added an explicit production `fb-system-detail` hide guard, and made the harness fail when either irrelevant block is visible. Static/contract, temporary-path persistence audit, 2025 focused IE 34/34, five-target build/Stage, performance, full `136 OK` harness, and 2019/2023/2025/2027 system-detail PNG generation passed. A new precise scan is needed for enriched metadata, and real wall/floor/roof assembly comparison remains queue item 27. | Needs Revit Check |
| Project Tracking / Element Ledger | Opt-in tracking, incoming reload exclusion, and recent history | `project-element-change-tracking/*`; `project-element-change-history`; App document lifecycle handlers; `FamilyBrowserElementChangeTrackingService` | Record only locally observed element changes after a successful Save/Save As/Sync, exclude incoming Reload Latest changes, preserve trustworthy baselines, and let an administrator inspect recent immutable history or explicitly export XLSX without creating automatic workbooks. | The first implementation had no dashboard reader for the stored ledger. Incoming Reload Latest activity could be attributed to the current workstation; a failed or partial initial baseline could later appear as mass-created history; a deletion first seen after baseline could be dropped; one lifecycle-service exception could prevent following services from committing; and baseline capture enumerated the document twice. | Added a Revit-version-compatible Reload Latest bridge with a transaction-name fallback for 2019, post-reload session rebase, fail-closed baseline creation with one-pass collection, incomplete previous-state deletion evidence, independently guarded lifecycle service calls, and a structured recent-history HTML viewer with explicit XLSX export only. Static/workflow/IE checks, five-target build/Stage, performance gate, ProgramData installation, and installed verification passed. Actual Revit event ordering and two-PC attribution remain queue item 40. | Needs Revit Check |

## Backups

| Time | Path | Reason |
|---|---|---|
| 2026-07-19 10:21 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\programdata-before-element-tracking-gap-audit-20260719-102125` | ProgramData checkpoint before deploying the element-tracking gap-audit build to all five Revit targets. |
| 2026-07-19 09:48 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\element-tracking-gap-audit-20260719-094812` | Source checkpoint before hardening Reload Latest attribution, baseline integrity, lifecycle isolation, and recent-history access. |
| 2026-07-15 16:50 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\system-detail-hidden-block-qa-20260715-165016` | Backup before hardening detached System detail against legacy Family/preview blocks, correcting the audit-only visibility override, and adding System detail PNG/DOM regression coverage. |
| 2026-07-15 16:08 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\system-layer-unit-ui-20260715-160848` | Backup of 26 existing source/audit files before adding structured compound-layer metadata, Revit-style layer rendering, synchronized persisted mm/in controls, and related static/IE automation. |
| 2026-07-15 15:30 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\family-browser-trade-switch-20260715-153052` | Backup before fixing Family/System ready-trade chip switching, prepared-slot/project-scan restoration, immediate DOM feedback, and its static/IE regression coverage. |
| 2026-07-13 15:39 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\pending-save-sync-lifecycle-20260713-153915` | Backup before adding provisional Family/System load/apply state, Save/Save As/Sync finalization, close-without-save discard, and the associated contract/UI harness coverage. |
| 2026-07-13 09:58 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\resizable-columns-numbered-pager-20260713-095810` | Backup before moving 150-row paging to the bottom status area, adding direct numbered pages, synchronized drag-resizable Family/System columns, and extending static/IE harness checks. |
| 2026-07-13 08:40 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260713\20260713-084004-live-mutation-refresh` | Backup before limiting startup preload to the initial shell render and adding persisted-state immediate-refresh regression guards. |
| 2026-07-13 08:18 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260713\20260713-081842-window-size-installer-ledger` | Audit-ledger backup before recording the final window-size installer, mail package, and integrity hashes. |
| 2026-07-11 10:37 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260711\20260711-103704-kky-tool-window-size` | Backup before matching the main Family Browser default/minimum window dimensions and working-area cap to KKY Tool, plus adding static regression guards. |
| 2026-07-10 15:57 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\structured-message-dialog-20260710-155755` | Backup before replacing the shared modern dialog plain-text body with structured HTML, adding copy/safe-close behavior, and extending static/WebBrowser message regression coverage. |
| 2026-07-10 15:43 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\managed-data-manual-refresh-invalidation-20260710-154314` | Incremental backup before routing manual homepage path/profile/URL/security refreshes through the shared prepared-data invalidation helper. |
| 2026-07-10 15:16 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\managed-data-read-audit-20260710-151626` | Backup before moving homepage bootstrap ahead of startup preload, repairing project scan alias persistence/read guards, adding stamp-safe alias matching, and adding managed-folder integrity automation. |
| 2026-07-10 14:29 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\registered-rvt-missing-list-empty-state-20260710-142951` | Backup before separating RVT-missing from RVT-connected/list-missing empty states, adding Family/System setup-state regression scenarios, correcting the audit fixture, and stabilizing queued-search validation. |
| 2026-07-10 14:02 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\admin-target-actions-layout-20260710-140211` | Backup before aligning Admin Standard active-target state, moving trade management beside the selector, normalizing Baseline/List action rows, and extending layout regression checks. |
| 2026-07-10 13:42 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\audit-target-scroll-preserve-20260710-134205` | Backup before preserving Model Check target-switch scroll state and adding WebBrowser scroll round-trip regression coverage. |
| 2026-07-10 11:15 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\family-browser-dom-virtualization-20260710-111520` | Backup before replacing all-row Family/System HTML with compact payloads, true 150-row DOM injection, cross-page selection adapters, and corresponding static/performance/WebBrowser harness checks. |
| 2026-07-10 10:49 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\family-browser-performance-shell-audit-20260710-104904` | Backup before separating startup-shell responsiveness from host assembly load/JIT in the performance harness and recording the final audit state. |
| 2026-07-10 10:13 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\family-browser-performance-gate-20260710-101349` | Backup before adding 2,000-row performance scenarios, cache timing, row-window/layout/language checks, and quality-gate integration. |
| 2026-07-10 09:33 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\family-browser-performance-v2-20260710-093316` | Backup before implementing V2 manifest/index/detail/thumbnail/project/row cache, asynchronous startup, lazy detail/preview loading, partial pane updates, and stale-source mutation guards. |
| 2026-07-09 23:17 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\include-2021-build-defaults-20260709-231700` | Backup before changing Family Browser quality-gate / UI harness defaults to include Revit 2021 and recording the updated deployment baseline. |
| 2026-07-09 17:00 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\scan-dialog-ok-cancel-xlsx-20260709-1700` | Backup before changing standard RVT precise-scan dialog OK/Cancel handling, adding auto-handled dialog action/button fields, exporting scan dialog diagnostics as XLSX, updating result labels, and extending static QA guards. |
| 2026-07-09 15:58 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\system-apply-selected-html-results-20260709-155842` | Backup before changing selected system-type multi-apply to selected-only preflight/post-check, replacing family/system confirmation/result dialogs with HTML operation dialogs, adding visible post-result refresh progress, and extending static QA guards. |
| 2026-07-09 15:08 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\system-type-detail-data-20260709-1508` | Backup before adding captured system-type detail summaries, routing/segment/dependency/layer table rendering, audit fixtures, and static/WebBrowser harness validation. |
| 2026-07-09 14:30 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\lookup-csv-detail-diff-20260709-143027` | Backup before adding lookup CSV row/column detail preview, lookup CSV fingerprint difference rows, audit fixtures, and static/WebBrowser harness validation. |
| 2026-07-09 14:12 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\debug-log-bottom-dock-20260709-141215` | Backup before removing the floating Debug Log FAB, docking the debug console at the bottom of the browser, and extending static/WebBrowser harness validation. |
| 2026-07-09 13:57 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\load-available-standard-detail-20260709-135746` | Backup before fixing `LoadAvailable` family detail to use standard snapshot content, suppress comparison diff sections, and extending static/WebBrowser harness validation. |
| 2026-07-09 13:28 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\search-focus-silent-detail-20260709-132803` | Backup before fixing Family/System Type live-search focus loss, adding quiet detail selection, and extending static/WebBrowser harness validation. |
| 2026-07-09 09:45 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\audit-target-selector-20260709-094538` | Backup before adding the Model Check review-target selector, fixing audit/admin action-grid class placement, and extending static/WebBrowser harness validation. |
| 2026-07-08 17:18 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\nested-child-dedupe-20260708-1718` | Backup before fixing composite nested-child dash-category duplicates, duplicate audit fixture rows, WebBrowser harness validation, and harness wrapper argument fallback. |
| 2026-07-08 16:57 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\detail-family-composition-preview-20260708-1645` | Backup before fixing detached detail family type table styling, type-parameter dropdown width/spacing, and small 3D preview fitting. |
| 2026-07-08 16:10 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\inline-status-filter-wrap-20260708-1610` | Backup before fixing the second Family/System Type action-row inline status filters, shell JS dynamic toolbar height, dense audit rows, and WebBrowser harness validation. |
| 2026-07-08 15:45 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\detail-diff-table-modal-20260708-1545` | Backup before fixing detail fingerprint diff concise summary, `상세 보기` button payload, detached diff modal, and WebBrowser harness validation. |
| 2026-07-08 15:05 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\admin-detail-preview-layout-20260708-1505` | Backup before fixing Admin Settings button layout, detached detail full-width parameters, composite nested-family table, and preview large-view behavior. |
| 2026-07-08 14:45 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\detail-auto-open-design-20260708-1445` | Backup before auto-opening detached detail on Family/System Type tab entry and refining detached detail window layout. |
| 2026-07-08 14:30 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\project-central-subtitle-tooltip-20260708-1430` | Backup before changing the header project subtitle to compact local/central file-name tokens with hover path tooltips. |
| 2026-07-08 14:20 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\filterbar-no-ellipsis-20260708-1420` | Backup before fixing Family/System Type status and discipline filter button clipping after scan. |
| 2026-07-08 14:00 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\admin-standard-buttons-layout-20260708-1400` | Backup before reorganizing the Admin Settings `기준 RVT` and `표시 표준 목록` button layout. |
| 2026-07-08 13:32 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\detail-preview-audit-20260708-133240` | Backup before adding detached detail 3D preview fallback and automated detail-content validation. |
| 2026-07-08 10:30 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\ui-audit-automation-20260708-103040` | Backup before adding automated UI audit contract, Revit-free render seam, WebBrowser click harness, and quality gate scripts. |
| 2026-06-29 15:03 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260629\20260629-150357-full-button-action-review` | Backup before first button/action review fixes. |
| 2026-06-29 15:29 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260629\20260629-1529-tracking-source-mismatch` | Backup before tracking source mismatch / dirty-marker refresh fix. |
| 2026-06-29 15:35 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260629\20260629-1535-existing-family-update-selection-key` | Backup before existing-family update selection key fix. |
| 2026-06-29 15:40 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260629\20260629-1540-debug-visibility-guard` | Backup before debug visibility / internal-path guard fix. |
| 2026-06-29 15:45 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260629\20260629-1545-modeler-family-update-ui-guard` | Backup before modeler family update-row UI/action guard fix. |
| 2026-06-29 15:52 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260629\20260629-1552-unregistered-request-actions` | Backup before unregistered family/system row request-conversion actions. |
| 2026-06-29 23:34 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260629\20260629-2334-admin-governance-controls` | Backup before attempted Admin Settings governance-control restoration. That source change was reverted after user clarified the tab split was intentional. |
| 2026-06-29 23:43 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups\20260629\20260629-2343-tab-scoped-governance-controls` | Backup before restoring governance controls to their split tabs. |
| 2026-06-30 14:53 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\standard-list-empty-action-20260630-145329` | Backup before adding the Admin Mode standard-list registration action to Family/System Type empty states. |
| 2026-06-30 15:11 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\button-response-fix-20260630-151103` | Backup before separating row selection from detached detail navigation and restoring list command/filter responsiveness. |
| 2026-06-30 15:20 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\permissions-file-guard-only-20260630-152007` | Backup before making the Permissions / Guard tab file-guard-only. |
| 2026-06-30 15:01 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\detail-window-open-fix-20260630-1515` | Backup before fixing detached detail-window open routing and JS scheduling. |
| 2026-06-30 15:30 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\ui-language-result-qa-20260630-153047` | Backup before language/result-window/text-safety QA fixes and static checker addition. |
| 2026-06-30 15:45 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\csv-difference-export-20260630-1545` | Backup before adding real CSV review export and system-type difference export columns. |
| 2026-06-30 15:56 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\family-size-table-lookup-fingerprint-20260630-155635` | Backup before fixing Revit family internal lookup CSV / size table fingerprint capture. |
| 2026-06-30 16:05 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\remove-review-csv-export-20260630-160542` | Backup before removing mistaken Current Model Check review-result CSV export and keeping `.xlsx` only. |
| 2026-06-30 16:36 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\unified-dialog-cleanup-20260630-163659` | Backup before routing dashboard, standard-family selection, and file-guard confirmation messages through the shared modern dialog. |
| 2026-07-06 15:57 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\audit-md-family-load-recheck-20260706-155716` | Backup before recording the selected family load recheck results in this audit MD. |
| 2026-07-06 16:05 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\browser-button-navigation-guard-20260706-160554` | Backup before fixing nested WebBrowser custom-scheme button navigation guards and static QA. |
| 2026-07-06 16:20 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\debug-log-visibility-fix-20260706-162012` | Backup before aligning Debug Log / F12 render and host-toggle conditions. |
| 2026-07-06 16:26 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\debug-log-f12-host-route-20260706-162628` | Backup before routing missing-panel F12 to the host `debug-log` action. |
| 2026-07-09 08:47 KST | `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups\detached-parameter-table-scroll-20260709-084750` | Backup before fixing detached detail parameter table width/overflow and long formula scroll containment. |

## Build / Install Log

| Time | Command | Result |
|---|---|---|
| 2026-07-15 17:14 KST | `Invoke-FamilyBrowserQualityGate.ps1 -Years 2019,2021,2023,2025,2027 -Configuration Release -OutputDir artifacts\family-browser-ui-audit\20260715-system-layer-unit-final-v2` | Passed static/contract and nested-family propagation, all five Release builds/Stage verification, 2,000-row performance/cache, and full IE HTML/click/language/layout/detail harness with `136 OK`, one expected Revit 2021 `SKIP runtime-not-installed`, and zero failures. System detail PNGs were generated for every installed runtime target and the 2025 full-height visual shows one routing/layer surface with no legacy empty preview or Family composition block. Performance: shell `3-13ms`, cold `452-783ms`, warm `304-444ms`, filter `11-17ms`. Stage SHA256: 2019/2021/2023 `86E469E8186B86D91D82BD38CF8F3656E09B9EA39B5A5C32D2874667D1BC4288`, 2025 `E70542DBC2D00AF883AD0C823ED672CB6F3D5778639E2AB25469ABB8453AF0EF`, 2027 `F98E659915870993F37A69894567CF7080C6421B65816DD916A923B203C6146D`. ProgramData/installers were not changed. |
| 2026-07-13 10:21 KST | `Invoke-FamilyBrowserQualityGate.ps1 -Years @('2019','2021','2023','2025','2027') -Install`; `Build-FamilyBrowserInstaller.ps1 -Version 1.0 -Label resizable-columns-numbered-pager-20260713 -MailPackageMinimumMB 15` | Passed static/contract, all five Release builds/stage verification, 2,000-row paging/selection/resize/cache gate, Korean/English IE click/layout/detail scenarios, ProgramData install, and installed verification with zero failures. Report: `artifacts\family-browser-ui-audit\20260713-100935-quality-gate`. Stage/ProgramData hashes match: 2019/2021/2023 `49785B35FAB6EA8B474E123B48843776006C8FA6A2EDC0CE07A26624D0B9E663`, 2025 `A0467536B434EF9AFE163AD45540BBAD82CB2D143670B8EA4F7F724D389C75AD`, 2027 `4F5D4F5709687ECEF0DF34FC3A77614F835244ACDA9D7819895C8AEF7E864C6B`. Installer SHA256 `BB361BDC2FC1301225426A63CFD1BCEC3BF2E13B30AB350FFA74FBA13E1976E8`; 15.5MB mail ZIP SHA256 `FC79AF599902F49C2AA57FFCC846E8205E785CF962B9763686BA20E0D6B63DF6`. |
| 2026-07-10 16:20 KST | `Invoke-FamilyBrowserQualityGate.ps1 -Years @('2019','2021','2023','2025','2027') -OutputDir artifacts\family-browser-ui-audit\structured-message-dialog-20260710-quality-gate -Install` | Passed static/contract, five-target Release build/stage, staged verification, 2,000-row performance/cache, HTML/click harness, ProgramData install, and installed verification. Harness: 89 results, 88 `OK`, Revit 2021 `SKIP runtime-not-installed`, zero failures; all eight Korean/English structured-message scenarios passed. Stage/installed hash prefixes matched: 2019/2021/2023 `E0A512DEFE22`, 2025 `B8DB19C7374A`, 2027 `5790392B58F2`. |
| 2026-07-10 15:46 KST | Final five-target install, installed verification, and stage-vs-installed SHA256 comparison | Installed/verified final manual-refresh invalidation build for all five ProgramData targets. Hash prefixes matched: 2019/2021/2023 `93844F5F0FC0`, 2025 `5FE64EB21D07`, 2027 `77900FEDE67A`. |
| 2026-07-10 15:45 KST | `Invoke-FamilyBrowserQualityGate.ps1 -SkipBuild -SkipHarness -ManagedDataAudit` after shared invalidation helper | Passed static/contract checks, five-target staged verification, and 2,000-row performance/cache gate. Managed path remained external-state `UNAVAILABLE`. Report: `artifacts\family-browser-ui-audit\managed-data-manual-refresh-20260710-gate`. |
| 2026-07-10 15:44 KST | Final `Build-FamilyBrowserRecovered.ps1` for 2019/2021/2023/2025/2027 | All five targets built and staged with zero errors after adding shared invalidation to deferred, manual path, profile, URL, security, and field-test routes. |
| 2026-07-10 15:40 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2021','2023','2025','2027')`; installed verification and stage-vs-installed SHA256 comparison | Installed and verified all five ProgramData targets. Hash prefixes matched: 2019/2021/2023 `76218D036EB2`, 2025 `AF2C8CB4CF58`, 2027 `F425DC73464E`. |
| 2026-07-10 15:39 KST | `Invoke-FamilyBrowserQualityGate.ps1 -Years @('2019','2021','2023','2025','2027') -SkipBuild -ManagedDataAudit` | Passed static/contract checks, staged verification, 2,000-row performance/cache gate, and 80 UI scenarios with zero failures; 2021 runtime was `SKIP runtime-not-installed`. Managed-data step reported external-state `UNAVAILABLE` because both homepage candidates were unreachable. Report: `artifacts\family-browser-ui-audit\managed-data-read-20260710-complete-quality-gate`. |
| 2026-07-10 15:30 KST | `Test-FamilyBrowserManagedData.ps1` live + fixture validation | Live homepage returned version `2026.05.19-kcim-test`, but `I:` and `D:\TEST` were unavailable and local real V2 source/row caches were empty. A complete D-drive fixture returned `OK`; removing its comparison report produced the expected integrity `FAIL`. Fixture was removed afterward. |
| 2026-07-10 15:24 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2021','2023','2025','2027') -Configuration Release`; staged verification | All five targets built and staged successfully. 2019/2021/2023 had zero errors; 2025/2027 retained only existing SDK/WindowsBase/platform warnings. |
| 2026-07-10 14:36-15:00 KST | `Invoke-FamilyBrowserQualityGate.ps1 -Years 2019,2021,2023,2025,2027 -Install`; corrected five-target build/stage; full UI harness; 2019 stable retry; ProgramData install/verify | Static/contract, five-target Release build/stage, staged verification, and 2,000-row performance/cache gate passed. The first full UI pass exposed that the old missing-RVT fixture incorrectly meant scan-needed; after correcting it, all new Family/System missing-list Korean/English scenarios passed on 2019/2023/2025/2027. One unrelated 2019 Korean queued-search timing check failed once, passed immediate 20/20 retry, was stabilized with a 2.5s render wait, and passed another 20/20 run. ProgramData installation and installed verification passed for all five targets; stage/installed DLL SHA256 matched. Revit 2021 runtime remained `SKIP runtime-not-installed`. |
| 2026-07-10 14:18 KST | `Invoke-FamilyBrowserQualityGate.ps1 -Years 2019,2021,2023,2025,2027 -OutputDir artifacts\family-browser-ui-audit\audit-scroll-admin-layout-20260710-complete-quality-gate` | Passed after Model Check scroll preservation and Admin Standard target/action layout fixes. Static/contract, five-target Release build/stage, staged verification, 2,000-row performance/cache, and UI harness passed. Harness: 73 results, 72 `OK`, Revit 2021 `SKIP runtime-not-installed`, zero failures. Revit 2019 was open, so ProgramData install was intentionally not run. |
| 2026-07-10 11:48 KST | `Invoke-FamilyBrowserQualityGate.ps1 -Years 2019,2021,2023,2025,2027 -OutputDir artifacts\family-browser-ui-audit\dom-virtualization-20260710-complete-quality-gate -Install`; `Build-FamilyBrowserInstaller.ps1 -Version 1.0 -Label dom-virtualization-20260710 -MailPackageMinimumMB 14.3` | Passed after true DOM virtualization. Static/contract, five-target build/stage, staged verification, 2,000-row performance/cache, Korean/English click/layout/detail harness, ProgramData install, and installed verification all passed. Runtime targets produced 1,000 total / 150 DOM / 150 visible rows and 9-10ms changing-query filters; 2021 runtime was `SKIP runtime-not-installed` while package/install checks passed. Installer SHA256 `FA190E3610B85BED52F2019958775655178BD44F67A1FFBDC86510E7D5F41D35`; 14.8MB mail ZIP SHA256 `E92FCF35CECD45C2E6CC407069345BDB82AF4FC6BD41C2EB37632A0116E2D26E`. |
| 2026-07-10 11:09 KST | `Invoke-FamilyBrowserQualityGate.ps1 -SkipBuild -OutputDir artifacts\family-browser-ui-audit\performance-v2-20260710-complete-quality-gate`; `Install-FamilyBrowserRecovered.ps1`; `Verify-FamilyBrowserRecovered.ps1 -Installed`; `Build-FamilyBrowserInstaller.ps1 -Version 1.0 -Label performance-v2-20260710` | Passed after V2 cache/async/lazy/performance implementation. Static/contract, 2019/2021/2023/2025/2027 stage, installed manifests/DLLs, Korean/English click/layout/detail scenarios, and 2,000-row performance gate passed. 2021 runtime smoke was skipped because Revit 2021 is not installed; package/install verification passed. Stage and installed DLL hashes matched for all five targets. Installer SHA256 `474EF374F8955345314A9240FC09BE169B289AE992A8ADBBD6FF7D1701C2C72D`; 13.6MB mail ZIP SHA256 `ED87385647C8E2BEC85BEC7667FA0093D13540DC4938EDC86ED34663AE54D16E`. |
| 2026-07-09 23:45 KST | `Invoke-FamilyBrowserQualityGate.ps1 -OutputDir artifacts\family-browser-ui-audit\include-2021-build-defaults-20260709-2338`; `Build-FamilyBrowserInstaller.ps1 -Version '1.0' -Label 'include-2021-defaults-20260709-2345'`; `Install-FamilyBrowserRecovered.ps1 -Years @('2021')`; `Verify-FamilyBrowserRecovered.ps1 -Installed` | Passed after making Revit 2021 part of the default Family Browser build / quality-gate flow. Static/contract checks passed, Release build/stage created 2019/2021/2023/2025/2027, staged addin verification passed for all five years, UI harness passed with 2021 runtime smoke marked `SKIP` because Revit 2021 is not installed on this PC, 2021 ProgramData addin install passed, and installed verification passed for 2019/2021/2023/2025/2027. Installer includes Rvt2021 files; exe SHA256 `7E646AD49C889E1B8B9EAB41C17298563C48EEE040982514174BE622B533F41F`, mail zip SHA256 `D1F4837A0204EF125C8579FBC13C6D4A6E6C333C1DBF6C0E75EAFC3731F4F61E`. |
| 2026-07-09 19:46 KST | `Invoke-FamilyBrowserQualityGate.ps1 -OutputDir artifacts\family-browser-ui-audit\scan-dialog-ok-cancel-xlsx-20260709-1700`; `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Passed after precise-scan dialog OK/Cancel/XLSX patch. Static/contract checks passed, Release build/stage passed with existing warnings only, staged add-in verification passed, WebBrowser harness passed 72 scenarios with 0 failures, ProgramData install completed for 2019/2023/2025/2027, and installed verification passed. Revit was not running during install. |
| 2026-07-09 16:16 KST | `Invoke-FamilyBrowserQualityGate.ps1 -OutputDir artifacts\family-browser-ui-audit\system-apply-selected-html-results-20260709-1610` plus direct 2019/2023/2025/2027 project builds during the patch | Passed after selected-system-apply and HTML operation result patch. UI static/contract checks passed, Release build/stage passed, staged add-in verification passed, and WinForms WebBrowser click harness passed 72 scenarios with 0 failures. This run did not install to ProgramData. |
| 2026-07-09 15:53 KST | `Test-FamilyBrowserUiStatic.ps1`; `Build-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027 -Configuration Release`; `Verify-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027`; `Invoke-FamilyBrowserUiAuditHarness.ps1 -OutputDir artifacts\family-browser-ui-audit\system-type-detail-data-20260709-1548`; `Install-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years 2019,2023,2025,2027` | Passed after system-type detail data patch. Static/contract checks passed, Release build completed with existing warnings only, stage verification passed, WebBrowser harness passed 72 scenarios with 0 failures, ProgramData install completed for 2019/2023/2025/2027, and installed verification passed. Revit was not running during install. |
| 2026-07-09 14:58 KST | `Test-FamilyBrowserUiStatic.ps1`; `Build-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027 -Configuration Release`; `Verify-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027`; `Invoke-FamilyBrowserUiAuditHarness.ps1 -OutputRoot artifacts\family-browser-ui-audit\lookup-csv-detail-diff-20260709-1452`; `Install-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years 2019,2023,2025,2027` | Passed after lookup CSV detail/difference patch. Static/contract checks passed, Release build completed with existing warnings only, stage verification passed, WebBrowser harness passed 72 scenarios with 0 failures, ProgramData install completed for 2019/2023/2025/2027, and installed verification passed. Revit was not running during install. |
| 2026-07-09 13:52 KST | `Test-FamilyBrowserUiStatic.ps1`; `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release`; `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')`; `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2019','2023','2025','2027') -OutputDir artifacts\family-browser-ui-audit\search-focus-silent-detail-20260709-1345`; `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Passed after the Family/System Type live-search quiet-detail patch. Static/contract checks passed, build completed with existing warnings only, staged verification passed, WebBrowser harness passed 72 scenarios with 0 failures, ProgramData install completed for 2019/2023/2025/2027, and installed verification passed. Revit was not running during install. |
| 2026-07-09 13:16 KST | `Build-FamilyBrowserInstaller.ps1 -Version '1.0' -Label 'model-check-target-20260709-1021' -MailPackageMinimumMB 13.1` | Built and stage-verified 2019/2021/2023/2025/2027, compiled installer exe, and created a 13.6 MB mail zip. Installer SHA256 `394C831D072E3F32C38A00D1E4DB64DA43745A7340D61097C72E69E86C07CAB9`; mail zip SHA256 `F47DC01392CADEADF309DA5C686831E9BA8B46DCCD7AA9B31292BAE2AE96A0F6`. Existing .NET/Revit reference warnings only; errors 0. |
| 2026-07-09 10:21 KST | `Test-FamilyBrowserUiStatic.ps1`; `Build-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027 -Configuration Release`; `Verify-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027`; `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years 2019,2023,2025,2027 -OutputDir artifacts\family-browser-ui-audit\audit-target-selector-20260709-1015`; `Install-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years 2019,2023,2025,2027` | Passed after adding the Model Check review-target selector and hardening the Model Check 2x2 action-card layout. Static/contract checks passed, build completed with existing warnings only, staged verification passed, WebBrowser harness passed 72 scenarios with 0 failures, ProgramData install completed for 2019/2023/2025/2027, and installed verification passed. |
| 2026-07-09 08:57 KST | Stage vs installed DLL SHA256 comparison | Installed ProgramData DLLs matched the current stage DLLs after the parameter-table patch: 2019/2023 `09D04891BB78`, 2025 `3DAA47BED9C0`, 2027 `62FB445250A0`. |
| 2026-07-09 08:56 KST | `Test-FamilyBrowserUiStatic.ps1`; `Build-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027 -Configuration Release`; `Verify-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027`; `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years 2019,2023,2025,2027 -OutputDir artifacts\family-browser-ui-audit\detached-parameter-table-scroll-20260709b`; `Install-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years 2019,2023,2025,2027` | Passed after detached detail parameter table width/overflow patch. Static/contract checks passed, build errors 0 with existing warnings, stage verification passed, WebBrowser harness passed all 32 scenario/year combinations, ProgramData install completed for 2019/2023/2025/2027, and installed verification passed. |
| 2026-07-09 08:24 KST | `ISCC.exe /DMyAppVersion=1.0 /DMyOutputBaseName=KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_nested-child-dedupe-20260709-0824_NoSetupLdr_Setup /DMyOutputDir=artifacts\family-browser\installers KKY_FamilyBrowser_Compiler_NoSetupLdr.iss` | Compiled a `UseSetupLdr=no` installer to work around PCs that block installer execution from `%TEMP%`. This is not a source-code change; it avoids the standard Inno setup-loader temp extraction path. Installer SHA256 `64F69F9996115532869B3B4F9685DCAA57584F31F718A6B717653FB3074F8059`; size 4.73 MB. |
| 2026-07-09 08:19 KST | `Build-FamilyBrowserInstaller.ps1 -Version '1.0' -Label 'nested-child-dedupe-20260709-0819-rebuild' -MailPackageMinimumMB 13.1` | Rebuilt a fresh installer and mail package from the same nested-child-dedupe source. Stage verification passed for 2019/2021/2023/2025/2027. Installer SHA256 `B6D9F8E915C44131F3D55C97E089C737B5EF691C813B7765B431C368D0D1A14A`; mail zip SHA256 `28A78C7CC9CED299FD88F2D8EAA92188FFDB1DD345C454BF990C254D743A0EEE`. Existing .NET/Revit reference warnings only; errors 0. |
| 2026-07-09 08:16 KST | `Build-FamilyBrowserInstaller.ps1 -Version '1.0' -Label 'nested-child-dedupe-20260709-0816' -MailPackageMinimumMB 13.1` | Built and stage-verified 2019/2021/2023/2025/2027, compiled installer exe, and created a 13.6 MB mail zip. Installer SHA256 `BD0722E76C96EDC1A376DDD02359D41AB6D8C3B9AA4CC0531C9D2C1D9F1AA80C`; mail zip SHA256 `52E16E59865C07D9EBDB66EDA764A45B548C91F7B11B55FCF62A3BECCBD62C16`. Existing .NET/Revit reference warnings only; errors 0. |
| 2026-07-08 17:23 KST | `Test-FamilyBrowserUiStatic.ps1`; `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release`; `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')`; `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2019','2023','2025','2027') -OutputDir artifacts\family-browser-ui-audit\nested-child-dedupe-20260708a`; `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Passed after composite nested-child dash-category dedupe patch. The harness wrapper needed a Windows PowerShell-compatible `Arguments` fallback, then passed all 32 scenario/year combinations. Build retained existing Revit/WinForms warnings only. Stage and installed verification passed for all 4 years. |
| 2026-07-08 17:08 KST | `Test-FamilyBrowserUiStatic.ps1`; `Build-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027 -Configuration Release`; `Verify-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027`; `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years 2019,2023,2025,2027 -OutputDir artifacts\family-browser-ui-audit\detail-composition-preview-20260708a`; `Install-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years 2019,2023,2025,2027` | Passed after detached detail family composition/type dropdown/3D preview fit patch. Build retained existing Revit/WinForms warnings only. Stage and installed verification passed for all 4 years; WebBrowser harness passed all 32 scenario/year combinations. |
| 2026-07-08 16:31 KST | `Build-FamilyBrowserInstaller.ps1 -Version '1.0' -Label 'inline-status-wrap-20260708-1630' -MailPackageMinimumMB 13.1` | Built and stage-verified 2019/2021/2023/2025/2027, compiled installer exe, and created a 13.6 MB mail zip. Installer SHA256 `09AAEDA170BA722CC0C450B3D47961335AAB3221723EF57CF73133E943C2F37F`; mail zip SHA256 `4E48B732E951EB74B7146DB6E94BAE01930780A980D493EAA82CE40B32DBCF40`. |
| 2026-07-08 16:24 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static and contract checks passed after adding action-row inline status filter wrap guards, shell JS dynamic-height guards, dense audit fixture rows, harness validation tokens, and a guard that rejects the legacy clipped `.inline-status-toggle` base CSS. |
| 2026-07-08 16:26 KST | `Build-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027 -Configuration Release` | Passed with existing warnings after the second-row inline status filter wrap patch. The build was invoked directly from the active PowerShell session so all four year stages were generated. |
| 2026-07-08 16:27 KST | `Verify-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027` | Stage verification passed for all 4 years. |
| 2026-07-08 16:29 KST | `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years 2019,2023,2025,2027 -OutputDir artifacts\family-browser-ui-audit\inline-status-filter-wrap-20260708b` | Passed for all 32 scenario/year combinations. The harness now fails if inline action status filters use ellipsis, clip visible text, or overlap the family/system grid. |
| 2026-07-08 16:30 KST | `Install-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years 2019,2023,2025,2027` | Installed to ProgramData and installed verification passed for all 4 years after the second-row inline status filter wrap patch. |
| 2026-07-08 15:55 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static and contract checks passed after adding fingerprint diff concise-summary guards, `data-diffraw` guards, detached diff modal guards, audit fixture diff rows, and harness validation tokens. |
| 2026-07-08 15:56 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings after the detail fingerprint diff table/modal patch. |
| 2026-07-08 15:57 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-07-08 16:03 KST | `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2019','2023','2025','2027') -OutputDir artifacts\family-browser-ui-audit\detail-fingerprint-diff-table-20260708a` | Passed for all 32 scenario/year combinations. The harness now checks concise diff summary text, clicks `상세 보기`, and verifies the fingerprint diff table modal contains Type Count and Width rows. |
| 2026-07-08 16:05 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed to ProgramData and installed verification passed for all 4 years after the detail fingerprint diff table/modal patch. |
| 2026-07-08 15:20 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static and contract checks passed after adding Admin Settings compact-grid guards, detached detail duplicate-title/full-width-parameter guards, nested child table guards, preview large-view clickable guard, and composite-family audit fixture guards. |
| 2026-07-08 15:27 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings after the Admin Settings/detail preview layout patch. |
| 2026-07-08 15:29 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-07-08 15:32 KST | `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2019','2023','2025','2027') -OutputDir artifacts\family-browser-ui-audit\admin-detail-preview-layout-20260708a` | Passed for all 32 scenario/year combinations. The harness now checks composite nested child table content, 3D preview markup, and `크게 보기` modal opening. |
| 2026-07-08 15:37 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed to ProgramData and installed verification passed for all 4 years after the Admin Settings/detail preview layout patch. |
| 2026-07-08 14:55 KST | `Build-FamilyBrowserInstaller.ps1 -Version '1.0' -Label 'detail-auto-open-20260708-1455' -MailPackageMinimumMB 13.1` | Built and stage-verified 2019/2021/2023/2025/2027, compiled installer exe, and created a 13.6 MB mail zip. Installer SHA256 `4618EC08CC73729211A26E2FCA73AE8015CB5B55F578FB04FF1416304D5ED677`; mail zip SHA256 `A79F93B2DE515D66EFDC8727233EDE3612B9BC854ABD2ACA2264DA30B6B868F0`. |
| 2026-07-08 14:45 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings after adding Family/System Type detached-detail auto-open and detached-detail layout refinements. |
| 2026-07-08 14:46 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years after the detached-detail auto-open/layout patch. |
| 2026-07-08 14:46 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static and contract checks passed after adding auto-open source guards, detached detail layout guards, and the `admin-system-with-data` audit scenario. |
| 2026-07-08 14:49 KST | `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2019','2023','2025','2027') -OutputDir artifacts\family-browser-ui-audit\detail-auto-open-design-20260708a` | Passed for all 32 scenario/year combinations. The harness now checks Family/System Type tab-entry emits `detail-window-open` when visible rows exist. |
| 2026-07-08 14:50 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed to ProgramData and installed verification passed for all 4 years after the detached-detail auto-open/layout patch. |
| 2026-07-08 14:29 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static and contract checks passed after adding compact local/central project subtitle guards. |
| 2026-07-08 14:30 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings after the project subtitle tooltip patch. |
| 2026-07-08 14:31 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-07-08 14:32 KST | `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2019','2023','2025','2027') -OutputDir artifacts\family-browser-ui-audit\project-central-subtitle-20260708a` | Passed for all 28 scenario/year combinations. The harness now checks local/central visible file names and hover title paths. |
| 2026-07-08 14:33 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed to ProgramData and installed verification passed for all 4 years. |
| 2026-07-08 14:12 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static and contract checks passed after adding filterbar no-ellipsis guards and harness tokens. |
| 2026-07-08 14:15 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings after Family/System Type filter wrapping patch. A wrapper invocation through `powershell -File` initially passed only 2019 due array argument handling, so the build was rerun by directly invoking the script from the active PowerShell session. |
| 2026-07-08 14:16 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years after direct script invocation. |
| 2026-07-08 14:19 KST | `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2019','2023','2025','2027') -OutputDir artifacts\family-browser-ui-audit\filterbar-no-ellipsis-20260708a` | Passed for all 28 scenario/year combinations. The harness now fails if rendered status/trade filter buttons still use ellipsis or clipped visible text. |
| 2026-07-08 14:20 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed to ProgramData and installed verification passed for all 4 years. |
| 2026-07-08 14:03 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static and contract checks passed after adding Admin Settings standard-action layout guards. |
| 2026-07-08 14:04 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings after Admin Settings button layout patch. |
| 2026-07-08 14:05 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-07-08 14:07 KST | `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2019','2023','2025','2027') -OutputDir artifacts\family-browser-ui-audit\admin-standard-buttons-layout-20260708b` | Passed for all 28 scenario/year combinations, including new `admin-standard-settings-layout`. |
| 2026-07-08 14:07 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | 2019 install failed because Revit 2019 is open and locking `KKY_FamilyBrowser_RevitHost.dll`. No source/build failure. |
| 2026-07-08 14:08 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2023','2025','2027')`; `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2023','2025','2027')` | 2023/2025/2027 installed to ProgramData and installed verification passed. 2019 remains pending until Revit 2019 closes. |
| 2026-07-08 13:35 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static and contract checks passed after adding detail-preview/source-token guards and correcting the contract token scope for harness-only `CheckBrowserDetailContent`. |
| 2026-07-08 13:36 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings after detached detail 3D preview fallback and audit scenario fixture changes. |
| 2026-07-08 13:37 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-07-08 13:43 KST | `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2019','2023','2025','2027') -OutputDir artifacts\family-browser-ui-audit\detail-preview-20260708a` | Passed for all 24 scenario/year combinations, including `admin-family-detail-with-preview`, which validates detail name/category/type composition/parameters/3D preview DOM. |
| 2026-07-08 13:44 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed updated Family Browser output to ProgramData addin folders for all 4 years. |
| 2026-07-08 13:45 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. Stage-vs-installed SHA256 hashes matched: 2019/2023 `666AA32F1CCD`, 2025 `CEF98154F5C8`, 2027 `8C77BA5C8042`. |
| 2026-07-08 10:55 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static and contract checks passed after adding the automated audit contract and render seam. |
| 2026-07-08 10:57 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings; stage created for 2019/2023/2025/2027. |
| 2026-07-08 10:58 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-07-08 11:25 KST | `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2019')` | First harness pass found the 2019/2023 generated-JS regex break and native DLL popup/timeout issue. |
| 2026-07-08 11:27 KST | `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2019')` | Passed after fixing the generated JS newline regex and harness hard-exit/native popup handling. |
| 2026-07-08 11:33 KST | `Invoke-FamilyBrowserUiAuditHarness.ps1 -Years @('2025','2027')` | Passed after adding Autodesk Shared RealDWG/Components dependency directories. |
| 2026-07-08 11:41 KST | `Invoke-FamilyBrowserQualityGate.ps1 -Years @('2019','2023','2025','2027') -SkipBuild` | Integrated quality gate passed: static/contract checks, stage reuse, staged addin verification, and WebBrowser click harness for all 4 years. Report: `artifacts\family-browser-ui-audit\quality-gate-20260708a`. |
| 2026-07-08 13:12 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed latest staged Family Browser output to ProgramData addin folders for all 4 years. |
| 2026-07-08 13:12 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. Stage-vs-installed SHA256 hashes matched: 2019/2023 `79022E95B75B`, 2025 `0F02BBCDA661`, 2027 `EC463CD4001E`. |
| 2026-07-08 13:17 KST | `Build-FamilyBrowserInstaller.ps1 -Version '1.0' -Label 'mail-20260708-1315' -MailPackageMinimumMB 13.1` | Built and stage-verified 2019/2021/2023/2025/2027, compiled installer exe, and created mail zip under 15 MB. Mail zip: `artifacts\family-browser\mail-packages\KKY_FamilyBrowser_RevitHost_2019_2021_2023_2025_2027_v1.0_mail_20260708.zip` (13.6 MB). |
| 2026-06-29 15:05 KST | `Build-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027 -Configuration Release` | Passed with existing warnings. |
| 2026-06-29 15:06 KST | `Verify-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027` | Stage verification passed. |
| 2026-06-29 15:07 KST | `Install-FamilyBrowserRecovered.ps1 -Years 2019,2023,2025,2027` | Installed to ProgramData addin folders. |
| 2026-06-29 15:07 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years 2019,2023,2025,2027` | Installed verification passed. |
| 2026-06-29 15:31 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed. 2019 had 0 warnings; 2025/2027 retained existing warnings. |
| 2026-06-29 15:31 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-29 15:32 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-29 15:32 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-29 15:37 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings. |
| 2026-06-29 15:37 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-29 15:38 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-29 15:38 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-29 15:42 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings. |
| 2026-06-29 15:43 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-29 15:43 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-29 15:43 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-29 15:47 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings. |
| 2026-06-29 15:47 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-29 15:47 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-29 15:47 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-29 15:56 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings. |
| 2026-06-29 15:57 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-29 15:57 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-29 15:57 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-29 16:10 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed. 2019/2023 had 0 warnings; 2025/2027 retained existing WindowsBase / WindowsDesktop SDK warnings. |
| 2026-06-29 16:10 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-29 16:10 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-29 16:11 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-29 23:46 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings. |
| 2026-06-29 23:46 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-29 23:47 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-29 23:47 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-30 14:55 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Passed with existing warnings. |
| 2026-06-30 14:55 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-30 14:56 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-30 14:57 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-30 15:03 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Passed with existing warnings. |
| 2026-06-30 15:03 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-30 15:04 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-30 15:04 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-30 15:16 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Passed with existing warnings. |
| 2026-06-30 15:16 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-30 15:16 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-30 15:17 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-30 15:21 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Passed with existing warnings. |
| 2026-06-30 15:21 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-30 15:21 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-30 15:22 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-30 15:34 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static checks passed. |
| 2026-06-30 15:35 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Passed with existing warnings after fixing `MethodInvoker` namespace ambiguity in 2025/2027. |
| 2026-06-30 15:36 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-30 15:36 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-30 15:37 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-30 15:46 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static checks passed, including review CSV writer branch and system-type difference export field checks. |
| 2026-06-30 15:47 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Passed with existing warnings after CSV/difference export patch. |
| 2026-06-30 15:47 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-30 15:48 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-30 15:48 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-30 15:50 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static checks passed after targeting the Current Model Check export dialog and rejecting accidental CSV on file-guard export. |
| 2026-06-30 15:51 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Passed with existing warnings after correcting the export filter location. |
| 2026-06-30 15:51 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-30 15:52 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-30 15:52 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-30 15:57 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static checks passed, including lookup CSV / family size table fingerprint regression checks. |
| 2026-06-30 15:59 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Passed with existing warnings after correcting direct PowerShell array invocation. Stage manifest contains 2019/2023/2025/2027. |
| 2026-06-30 15:59 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-30 16:00 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-30 16:00 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-30 16:08 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static checks passed after making Current Model Check review export `.xlsx` only and rejecting review-result CSV writer/filter. |
| 2026-06-30 16:10 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Passed with existing warnings after removing review-result CSV export. |
| 2026-06-30 16:10 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-30 16:10 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-30 16:10 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-06-30 16:38 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static checks passed, including shared modern message dialog routes and raw dashboard/selection/file-guard MessageBox regression guards. |
| 2026-06-30 16:39 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings after adding `FamilyBrowserModernMessageDialog` and routing dashboard/selection/file-guard messages through it. |
| 2026-06-30 16:40 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-06-30 16:41 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-06-30 16:42 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-07-06 15:55 KST | Selected family load path static recheck | UI/router/handler/execution-service tokens present in all 3 host projects; execution path reaches `familyDoc.LoadFamily(targetDocument, loadOptions)`. |
| 2026-07-06 15:55 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static checks passed. |
| 2026-07-06 15:56 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed. 2019/2023 had 0 warnings; 2025/2027 retained existing WindowsDesktop SDK / WindowsBase warnings. |
| 2026-07-06 15:56 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-07-06 15:56 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-07-06 15:56 KST | Stage vs installed DLL SHA256 comparison | Installed ProgramData DLLs matched the current stage DLLs for 2019, 2023, 2025, and 2027. |
| 2026-07-06 16:07 KST | Dashboard `kkyfb:` action extraction | 77 generated action tokens checked in each host dashboard file; unknown action count was 0 for 2019-2023, 2025, and 2027. |
| 2026-07-06 16:07 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static checks passed, including nested WebBrowser `about:` custom-scheme navigation guards. |
| 2026-07-06 16:08 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings after browser-button navigation guard patch. |
| 2026-07-06 16:09 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-07-06 16:09 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-07-06 16:10 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-07-06 16:10 KST | Stage vs installed DLL SHA256 comparison | Installed ProgramData DLLs matched the current stage DLLs for 2019, 2023, 2025, and 2027. |
| 2026-07-06 16:20 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static checks passed, including Debug Log / F12 panel-gate regression checks. |
| 2026-07-06 16:21 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings after Debug Log / F12 fix. |
| 2026-07-06 16:22 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-07-06 16:22 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-07-06 16:23 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-07-06 16:24 KST | Stage vs installed DLL SHA256 comparison | Installed ProgramData DLLs matched the current stage DLLs for 2019, 2023, 2025, and 2027. |
| 2026-07-06 16:26 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static checks passed, including missing-panel F12 host-route regression checks. |
| 2026-07-06 16:27 KST | `Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release` | Passed with existing warnings after missing-panel F12 host-route fix. |
| 2026-07-06 16:27 KST | `Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Stage verification passed for all 4 years. |
| 2026-07-06 16:27 KST | `Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')` | Installed to ProgramData addin folders for all 4 years. |
| 2026-07-06 16:28 KST | `Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')` | Installed verification passed for all 4 years. |
| 2026-07-06 16:28 KST | Stage vs installed DLL SHA256 comparison | Installed ProgramData DLLs matched the current stage DLLs for 2019, 2023, 2025, and 2027. |
| 2026-07-06 16:29 KST | `Test-FamilyBrowserUiStatic.ps1` | UI static checks passed after clarifying the debug QA assertion label. |

| 2026-07-10 23:17 KST | Direct host builds: 2019-2023 / 2025 / 2027 | Passed after shared dialog/subwindow/native WinForms theme integration. Existing decompiled-source and Windows-platform warnings remain; zero compile errors. |
| 2026-07-10 23:18 KST | `Test-FamilyBrowserUiStatic.ps1` + `Test-FamilyBrowserUiContract.ps1` | Passed. All three hosts expose 78 generated actions, 241 exact routes, 50 prefix routes, and 11 browser-only functions; `theme-toggle` is contract-required. |
| 2026-07-11 00:54 KST | KKY Tool theme focused viewport gate | `viewport-1280-family` passed in KO/EN x light/dark. Fixed the audit-only `WebBrowser.Refresh()` page reload that had caused Korean screenshot scenarios to time out after a valid PNG was written. |
| 2026-07-11 01:24 KST | Full five-target quality gate | Static/contract, five-target stage verification, and 2,000-row performance/cache checks passed. All product DOM, route, click, language, theme, detail, and viewport assertions passed; 16 rows were marked failed only because long scenario paths prevented the auxiliary Edge PNG from being written. |
| 2026-07-11 01:31 KST | Screenshot-path regression rerun | Moved Edge profile/output work files to short `%TEMP%` paths. The 48 previously affected KO/EN x light/dark detail/admin combinations passed across 2019/2023/2025/2027; Revit 2021 recorded `SKIP runtime-not-installed`. |
| 2026-07-11 01:36 KST | ProgramData install and verification | Installed and verified 2019/2021/2023/2025/2027. Stage and installed DLL SHA256 values matched for all five targets. |
| 2026-07-11 01:43 KST | v1.0 installer/mail package | Inno Setup compiled the five-target installer. Final installer SHA256 `F48266A6AE57F01EC73531317C3B0CC383837E3A0B7B2D0E80A54CDD1F0CF855`; 15.25 MB mail package SHA256 `A1239BBBEC1BA32F6F237FC82435B12EA43A890738A3C8C5EB9365520B65B172`. |
| 2026-07-11 01:48 KST | Final visual set | Home, 1280 Family, Standard Management, detached detail, and structured-message KO/EN x light/dark scenarios passed 20/20 with final images under `artifacts/family-browser-ui-audit/20260711-kky-tool-theme-final-visuals`. |
| 2026-07-11 02:07 KST | Fixed-palette source checkpoint | Backed up the shared theme service, all three dashboard hosts, contract, wrapper, harness, and audit ledger to `artifacts/source-checkpoints/20260711-020736-family-browser-light-only`. |
| 2026-07-11 02:11 KST | Fixed-palette static/contract gate | Passed for all three hosts with 77 generated actions, 240 exact routes, 50 prefix routes, and 11 browser-only functions. `kkyfb:theme-toggle` is forbidden and no visible theme action remains. Evidence: `artifacts/family-browser-ui-audit/20260711-light-only-contract`. |
| 2026-07-11 02:13 KST | Focused fixed-palette visual/click gate | Home, Family detail, Standard Management, and 1280 Family Korean/English scenarios passed 10/10. Header exposes Refresh/Admin/Language only. Evidence: `artifacts/family-browser-ui-audit/20260711-light-only-focus`. |
| 2026-07-11 02:24 KST | Full five-target fixed-palette quality gate | Static/contract, 2019/2021/2023/2025/2027 build/stage, staged verification, 2,000-row performance/cache, single-palette KO/EN DOM/click/layout/detail/message checks, ProgramData install, and installed verification passed. Harness: 113 results, 112 `OK`, Revit 2021 `SKIP runtime-not-installed`, zero failures. Evidence: `artifacts/family-browser-ui-audit/20260711-light-only-quality-gate`. |
| 2026-07-11 02:28 KST | Fixed-palette installer/mail package | Installer SHA256 `48520F1ACA97BAB4B86F3C3BDF6A6BF45A232DC28355B5F256DEFFD074A8BCB4`; 15.25 MB mail package SHA256 `8FFE386E4784E2D79431705CAA00FA9EF126EB652ACC2641532D9681F48985A3`. |
| 2026-07-11 02:30 KST | Final stage-vs-installed SHA256 check | ProgramData DLLs exactly match current stage for all five targets. Hashes: 2019/2021/2023 `C390749BBFD052CCD5D27EFC0AAA04DD8DDC167F088D9530091BCBC3D55102D2`, 2025 `66F7C2CC7B416E2A2F7D01BA07EEE2B13B5938824723E7DF8C648D5606331BBF`, 2027 `134E541718E262752583FE3D639E3DB77A68B8FA732B44261EBBB17604F919DE`. |
| 2026-07-11 09:51 KST | Runtime-language source checkpoint | Backed up all three dashboard/scenario files, harness, wrapper, contract, static checker, and audit ledger to `artifacts/source-checkpoints/20260711-095147-language-toggle-purity`. |
| 2026-07-11 09:58 KST | Strengthened language static/contract gate | Passed for all three hosts with 77 generated actions, 247 exact routes, 51 prefix routes, and 11 browser-only functions. Required tokens now include transition initialization, cached-state translators, semantic discipline labels, and window-title diagnostics. |
| 2026-07-11 10:05 KST | First 2025 Korean -> English transition matrix | The strengthened harness correctly exposed three English data-row failures caused by partially translated Korean audit notes. Header status and unsaved-project transition scenarios already passed. Added whole-sentence row-note translations rather than allowing the Korean snippets. |
| 2026-07-11 10:09 KST | Focused 2025 transition rerun | All 28 Korean/English structured-message, Home, Family, System, missing-RVT/list, Standard Management, Model Check, requests, and viewport scenarios passed. Evidence: `artifacts/family-browser-ui-audit/20260711-language-toggle-focus-2025-rerun`. |
| 2026-07-11 10:25 KST | Full five-target runtime-language quality gate | Static/contract, 2019/2021/2023/2025/2027 build/stage, staged verification, 2,000-row performance/cache, IE DOM/click/language/layout/detail checks, ProgramData install, and installed verification passed. Harness: 113 results, 112 `OK`, Revit 2021 `SKIP runtime-not-installed`, zero failures. All 56 English transition results passed with zero language leakage. Evidence: `artifacts/family-browser-ui-audit/20260711-language-toggle-quality-gate`. |
| 2026-07-11 10:28 KST | Runtime-language installer/mail package and final hashes | Installer SHA256 `BE8495AF2EBE2CF30AAEB4264EE40AA626DD9F7D3A76302736DACBA78EEC9FFF`; 15.2 MB mail package SHA256 `0E1B9E06A37440878E455AE39A15A41278D9ED89BD4EBE1FD8C6464A84011BC9`. Stage/ProgramData DLL hashes match: 2019/2021/2023 `1F8FB7C72FE3BE6E2F3E11715409341A500CBDB27ED11726F13C2ED3D1E70BA3`, 2025 `DECF7E1E6776EC6ABE399BF89B0DA80179EE7895FC28CA430DC87BA12CBCC5DF`, 2027 `36FEA24E556C4DBD2E7EE8A8E9D5FBB4EACA557735B40F10C5825530345BCA99`. |
| 2026-07-11 10:51 KST | Full five-target KKY Tool-matched window-size quality gate | Static/contract, 2019/2021/2023/2025/2027 Release build/stage, staged verification, 2,000-row performance/cache, IE DOM/click/language/layout/detail checks, ProgramData install, and installed verification passed. Harness: 112 `OK`, Revit 2021 one `SKIP runtime-not-installed`, zero failures. Evidence: `artifacts/family-browser-ui-audit/20260711-103903-kky-tool-window-size`. |
| 2026-07-11 10:52 KST | Window-size deployment stage-vs-installed SHA256 | ProgramData DLLs exactly match stage: 2019/2021/2023 `CC47F8C2B0926424F1EF8DE81F8F312BB300E4FBF53E94C1022A9E66B57DA233`, 2025 `BB43A15EA9C127DC6A328EF220F3AA31DA7F0B5FA07339F0F74A242337B6D700`, 2027 `3235630E5BE8D8559A9D491861BA5AB9DA86818EDFE0750F22FF6579C81CC6BA`. |
| 2026-07-13 08:18 KST | Window-size installer preflight | Static/contract checks passed for all three hosts with 77 generated actions, 247 exact routes, 51 prefix routes, and 11 browser-only functions. Release build/stage and stage verification passed for 2019/2021/2023/2025/2027; zero compile errors. |
| 2026-07-13 08:18 KST | Window-size v1.0 installer/mail package | Inno Setup created `KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_kky-tool-window-size-20260713_Setup.exe` (3.32 MB), SHA256 `0426F791F478F74877B6123DDE678F562200D00DBDE203BA376AE4E6C0F46C74`. Mail ZIP `20260713_01.zip` is 15.5 MB, SHA256 `5D46080164B8098257B606044FCAA8B71E786FD8FAE1953D44EED626825CFBED`; its embedded `Setup.exe` hash matches the standalone installer. |
| 2026-07-13 08:56 KST | Full five-target live-state refresh quality gate | Static/contract including 18 mutation-refresh methods, 2019/2021/2023/2025/2027 build/stage, staged verification, 2,000-row performance/cache, IE DOM/click/language/layout/detail checks, ProgramData install, and installed verification passed. Harness: 112 `OK`, Revit 2021 one `SKIP runtime-not-installed`, zero failures. Evidence: `artifacts/family-browser-ui-audit/20260713-live-mutation-refresh-quality-gate`. |
| 2026-07-13 08:57 KST | Live-state stage-vs-installed SHA256 | ProgramData DLLs exactly match stage: 2019/2021/2023 `59985749B7BCB45BC3B81E937025F8023F9CAD470DC00FE179B380A8B67E50BD`, 2025 `373966B952B820B6E707CCEDA11E140D23F5BBF30766AC96556345F35479E9B5`, 2027 `BD9B3D7DA84AAC891F0CA0D3046A3A23A3E1CE104CEB77443B9B36BBDB9D7276`. |
| 2026-07-13 08:57 KST | Live-state refresh installer/mail package | Installer `KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_live-state-refresh-20260713_Setup.exe` SHA256 `F7DFF5EEAB7FC46D26A72846E9EB08F8290EEAED7D1BAE2B5643860A4EE986CA`; mail ZIP `20260713_02.zip` is 15.5 MB, SHA256 `688D20BCD2D7899F883F9476B73D62D062EEF9A5CCA1196F1E00EABE4A1B37FC`. Embedded `Setup.exe` hash matches the standalone installer. |

## 2026-07-13 Persisted-State Live Refresh

### Problem

- Trade rename and other policy changes were persisted correctly, but the immediate shell refresh could render `_startupPreloadResult` from the first browser open.
- The old startup policy or prepared slot metadata replaced the just-saved in-memory/live-store state, so users saw the change only after pressing Refresh.

### Changes

- Added `allowStartupPreload` to the shared shell refresh path. Only `CompleteInitialOpenRefresh(...)` passes `true`; every later refresh defaults to `false`.
- Applied the startup permission across the complete shell rebuild, including nested prepared-slot lookups and row-cache key generation, then reset it in `finally`.
- Normal refreshes now load `FamilyBrowserStandardPolicyStore` and current registration/list metadata while retaining existing tab, filter, selected-row, and scroll restoration.
- Added runtime source diagnostics and static guards for preload scope, live-store fallback, prepared-slot gating, and cleanup.
- Added a persisted-state refresh contract covering 18 state-writing methods, including the reported trade rename path.

### Verification

- Source checkpoint: `artifacts/source-file-backups/20260713/20260713-084004-live-mutation-refresh`.
- Full gate: `artifacts/family-browser-ui-audit/20260713-live-mutation-refresh-quality-gate` (112 `OK`, one Revit 2021 runtime skip, zero failures).
- Build/stage/install: 2019/2021/2023/2025/2027 passed; ProgramData DLL hashes match stage exactly.
- Installer: `artifacts/family-browser/installers/KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_live-state-refresh-20260713_Setup.exe`, SHA256 `F7DFF5EEAB7FC46D26A72846E9EB08F8290EEAED7D1BAE2B5643860A4EE986CA`.
- Mail package: `artifacts/family-browser/mail-packages/20260713_02.zip`, 15.5 MB, SHA256 `688D20BCD2D7899F883F9476B73D62D062EEF9A5CCA1196F1E00EABE4A1B37FC`.
- Needs Revit Check: on the external test PC, rename one trade and confirm the current target text, pressed target chip, header standard pill, and all trade selectors update immediately without Refresh.

## 2026-07-11 KKY Tool-Matched Default Window Size

### Problem

- KKY Tool opens at a desired `1400 x 900`, capped to 93% of the working area with a practical `1100 x 720` minimum.
- Family Browser had an independent wide-screen sizing formula that opened up to `1680 x 960` or `1780 x 1040`, so the two tools did not feel like the same product even after their visual styling was aligned.

### Changes

- Replaced `GetRecommendedStartSize(...)` in the 2019-2023, 2025, and 2027 hosts with the KKY Tool baseline: width `min(1400, workingWidth * 0.93)` and height `min(900, workingHeight * 0.93)`.
- Replaced the adaptive minimum floors with `min(1100, startupWidth)` and `min(720, startupHeight)`, matching KKY Tool while remaining valid on small work areas.
- Preserved the existing Revit-main-window monitor resolver, explicit working-area centering, normal startup state, and free user resizing.
- Added static source guards so any future host drift in desired size, working-area cap, minimum size, or centering fails the quality gate.

### Verification

- Source checkpoint: `artifacts/source-file-backups/20260711/20260711-103704-kky-tool-window-size`.
- Full gate evidence: `artifacts/family-browser-ui-audit/20260711-103903-kky-tool-window-size` (112 `OK`, one Revit 2021 runtime skip, zero failures).
- Build/stage/install: 2019/2021/2023/2025/2027 passed; ProgramData DLL hashes match stage exactly.
- Installer: `artifacts/family-browser/installers/KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_kky-tool-window-size-20260713_Setup.exe`, SHA256 `0426F791F478F74877B6123DDE678F562200D00DBDE203BA376AE4E6C0F46C74`.
- Mail package: `artifacts/family-browser/mail-packages/20260713_01.zip`, 15.5 MB, SHA256 `5D46080164B8098257B606044FCAA8B71E786FD8FAE1953D44EED626825CFBED`.
- Needs Revit Check: open Family Browser once at the workstation DPI and confirm its initial outer dimensions visually match KKY Tool and remain centered on the Revit monitor.

## 2026-07-11 KKY Tool Theme Redesign (Historical)

> Superseded at 2026-07-11 02:07 KST by the fixed single-palette decision below. The light/dark implementation details in this section are retained only as change history and are not current product behavior.

### Problem

- Family Browser still used the previous green brand language while KKY Tool used navy/blue, and the three Revit hosts duplicated shell CSS/JS.
- Main, detail, result, selection, RVT manager, Guard, sheet, and request windows did not share one theme lifecycle.
- The prior screenshot gate could reload the page while capturing and used deeply nested Edge profile paths, producing false timeout/screenshot failures.

### Changes

- Added shared IE-safe static theme assets under `KKY_FamilyBrowser_SharedUi`, including the compact KKY logo, light/dark CSS, theme JS, and common dialog/subwindow styling. Host-specific shell assets now consume the shared source instead of maintaining separate visual definitions.
- Added `FamilyBrowserUiTheme` (`Light`, `Dark`), `FamilyBrowserUiThemeService`, default `light`, and persisted `theme.txt` in the existing user settings area. Save failure keeps the session theme and logs diagnostics.
- Added required `kkyfb:theme-toggle` routing. Theme switching changes only `theme-light`/`theme-dark` and `data-theme`; it does not replace `DocumentText`, so active tab, search, filters, checked rows, selected detail, and scroll are preserved.
- Rebuilt the product header around the KKY logo/name, right-aligned document context and command buttons, wrapped status row, navy workflow rail, blue/cyan active controls, and neutral success/warning/error semantics.
- Applied the same theme to Home, Family, System Type, Requests, Model Check, Standard Management, detached detail/3D preview, parameter/Fingerprint/system tables, structured result/error dialogs, selection/RVT manager/File Guard/sheet windows, and input-heavy native forms.
- Removed the unused legacy `DashboardMessageDialog` and `FamilyLoadResultDialog` implementations after confirming the shared HTML result path is the active route.
- Extended the contract/harness for KO/EN x light/dark, theme state preservation/timing, brand lockup, contrast, previous-green regression, detached-detail duplicate title, and `1280x720`, `1600x900`, `1920x1080` layout checks.
- Fixed audit capture reliability: replaced navigation-causing `WebBrowser.Refresh()` with control repaint only, exported a script-free rendered DOM for visual capture, and used short `%TEMP%` Edge profile/PNG paths before copying artifacts to the report directory.

### Verification

- Source checkpoint: `artifacts/source-checkpoints/20260710-223820-kky-tool-theme-redesign`.
- Static/contract: passed; each host reports 78 generated actions, 241 exact routes, 50 prefix routes, and 11 browser-only functions.
- Performance/cache: 2,000-row gate passed with the existing 150-row windowing/state-preservation checks.
- Full matrix evidence: `artifacts/family-browser-ui-audit/20260711-kky-tool-theme-final-quality-gate`.
- Screenshot-path fix evidence: `artifacts/family-browser-ui-audit/20260711-kky-tool-theme-screenshot-recheck` (48 `OK`, Revit 2021 `SKIP runtime-not-installed`, zero failures).
- Latest Standard Management visual smoke: `artifacts/family-browser-ui-audit/20260711-kky-tool-theme-visual-export-smoke`.
- Consolidated final visuals: `artifacts/family-browser-ui-audit/20260711-kky-tool-theme-final-visuals` (20 `OK`, zero failures).
- Build/stage/install: 2019/2021/2023/2025/2027 passed; installed ProgramData DLL hashes match stage.
- Installer: `artifacts/family-browser/installers/KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_kky-tool-theme-20260711_Setup.exe` (3.32 MB).
- Mail package: `artifacts/family-browser/mail-packages/20260711_02.zip` (15.25 MB).
- Superseded runtime item: light/dark switching and `theme.txt` persistence no longer apply. Current runtime check is limited to confirming that no theme control appears and that all main/detail/result windows retain the fixed palette at the workstation DPI. Revit 2021 runtime remains unavailable on this PC.

## 2026-07-11 Fixed Single-Palette Follow-up

### Problem

- The user chose one consistent Family Browser appearance and did not want a theme setting or visible light/dark wording.
- The previous theme button, host route, persistence file, and two-theme audit matrix were now unnecessary product surface and test complexity.

### Changes

- Removed the header theme button, `theme-toggle` UI-only route, action handler, and dashboard theme setter from the 2019-2023, 2025, and 2027 hosts.
- Changed `FamilyBrowserUiThemeService` to always return the fixed default palette, ignore stale `theme.txt`, and perform no preference writes.
- Removed theme switching from the contract and added `kkyfb:theme-toggle` as a forbidden source token.
- Reduced the harness and scenario wrapper to one palette while preserving Korean/English, contrast, old-green regression, layout, click, message, detail, and performance checks.
- Kept dormant alternate-theme CSS/interfaces as non-routed compatibility internals to avoid an unrelated broad refactor; no product control, action, or persisted setting can activate them.

### Verification

- Source checkpoint: `artifacts/source-checkpoints/20260711-020736-family-browser-light-only`.
- Static/contract evidence: `artifacts/family-browser-ui-audit/20260711-light-only-contract` (77 actions, 240 exact routes, 50 prefix routes, 11 browser-only functions per host).
- Focused visual evidence: `artifacts/family-browser-ui-audit/20260711-light-only-focus` (10/10 `OK`).
- Full gate evidence: `artifacts/family-browser-ui-audit/20260711-light-only-quality-gate` (112 `OK`, Revit 2021 one runtime skip, zero failures).
- Build/stage/install: 2019/2021/2023/2025/2027 passed; current stage and ProgramData DLL hashes match exactly.
- Installer: `artifacts/family-browser/installers/KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_light-only-20260711_Setup.exe`, SHA256 `48520F1ACA97BAB4B86F3C3BDF6A6BF45A232DC28355B5F256DEFFD074A8BCB4`.
- Mail package: `artifacts/family-browser/mail-packages/20260711_03.zip`, 15.25 MB, SHA256 `8FFE386E4784E2D79431705CAA00FA9EF126EB652ACC2641532D9681F48985A3`.

## 2026-07-11 Runtime Language Transition Purity

### Problem

- English rendered cleanly when selected before dashboard state was created, but switching an already-open Korean dashboard to English left cached Korean strings visible.
- The leaked fields included the standard summary trade/count words, permission/tracking pills, unsaved-project placeholder, native window title, stored Family/System discipline labels, and some row notes.
- The previous language gate only tested direct-language rendering, so it could not reproduce this lifecycle bug.

### Changes

- Added bidirectional translators for cached standard, permission, tracking, project-title, and project-path placeholders; `SetLanguage` now updates the native form title immediately.
- Family/System row payloads, tree labels, legacy row markup, and modeler rows now resolve display discipline from the semantic discipline key in the current language instead of reusing a stored localized label.
- Added `InitialLanguageCode` to the Revit-free audit scenario. Every English dashboard scenario now builds Korean cached state first and then switches to English before HTML generation.
- Added an unsaved-project transition fixture and `data-window-title` audit state. The harness checks translated project subtitle and window title in addition to all visible text snippets.
- Kept `한국어` as the only permitted Hangul in English mode because it is the intentional label of the language-switch target button.

### Verification

- Source checkpoint: `artifacts/source-checkpoints/20260711-095147-language-toggle-purity`.
- Focused 2025 evidence: `artifacts/family-browser-ui-audit/20260711-language-toggle-focus-2025-rerun` (28/28 `OK`).
- Full evidence: `artifacts/family-browser-ui-audit/20260711-language-toggle-quality-gate` (112 `OK`, one Revit 2021 runtime skip, zero failures; 56 English transition results, zero language failures).
- Build/stage/install: 2019/2021/2023/2025/2027 passed and ProgramData hashes match stage.
- Installer: `artifacts/family-browser/installers/KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_language-toggle-purity-20260711_Setup.exe`, SHA256 `BE8495AF2EBE2CF30AAEB4264EE40AA626DD9F7D3A76302736DACBA78EEC9FFF`.
- Mail package: `artifacts/family-browser/mail-packages/20260711_04.zip`, 15.2 MB, SHA256 `0E1B9E06A37440878E455AE39A15A41278D9ED89BD4EBE1FD8C6464A84011BC9`.

## 2026-07-13 Full HTML Dialog Shell

### Problem

- Scan/result/error contents were HTML, but the dialog caption, subtype, close action, footer hint, and OK/Yes/No/copy buttons were still WinForms controls.
- The mixed HTML/WinForms composition used separate DPI/font measurement paths, so titles and action labels could clip or render incorrectly even when the body looked correct.
- Family/System load/apply result dialogs had the same partial conversion: the result table was HTML while the outer form footer remained WinForms.

### Changes

- Added `KKY_FamilyBrowser_SharedUi\FamilyBrowserHtmlDialogHost.cs` as the single visible dialog host for all target versions.
- The native form is borderless and contains only one full-dock `WebBrowser`; title, subtitle, severity accent, close/maximize actions, scrollable content, status, copy/open-folder/copy-path, and OK/Yes/No actions are one HTML document.
- Routed shared modern messages and all three host copies of `FamilyBrowserOperationHtmlDialog` through the new host, removing their visible WinForms `Label`, `Button`, `TableLayoutPanel`, and footer shells.
- Preserved close safety and keyboard behavior: closing/Esc on Yes/No returns `No`, Enter follows the configured default, and OK remains `DialogResult.OK`.
- Updated static QA and the message IE harness to require the full HTML shell semantic IDs and reject reintroduction of WinForms title/action controls.

### Verification

- Backup: `_backups\full-html-dialog-shell-20260713-092523`.
- Focused 2025 IE evidence: `artifacts\family-browser-ui-audit\full-html-dialog-smoke-20260713`; 28/28 scenarios passed, and both Korean/English message scenarios rendered three expected HTML actions with zero failures.
- Full evidence: `artifacts\family-browser-ui-audit\20260713-full-html-dialog-quality-gate`; static/contract, five-target build/stage, Stage verification, 2,000-row performance/cache, IE DOM/click/language/layout/detail, ProgramData install, and installed verification all passed. Harness result: 112 `OK`, one Revit 2021 runtime skip, zero failures.
- Stage/ProgramData SHA256: 2019/2021/2023 `5AD15DF6D84EE345F0A2E299A8B3313A1F5016041375B8639CD9C5C301EC3C4A`; 2025 `6D3ED7DE304A69B4EC86D2EAD4E683B243F3D1570A458DE5F85569919E0A1348`; 2027 `2AFCB536CC81B16118D15DFEAB10233B4F21473C412B14F910EBE712589C8A66`.
- Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_full-html-dialog-20260713_Setup.exe`, SHA256 `2DFAB5B62D41E49532FF90680FD0BC23C64D1EED275E07BB9148ED89BF8D47B8`.
- Mail package: `artifacts\family-browser\mail-packages\20260713_03.zip`, 15.5 MB, SHA256 `28F4F5381F11281239CF5FA690585CB806B2ACFEF2D7BF29AEE217B77F66B85E`.
- `Needs Revit Check`: trigger a real precise-scan completion/error and Family/System apply result at the workstation DPI, verify the full title/footer labels, drag/maximize, copy, report-folder/path, OK/Yes/No, Enter, Esc, and close actions visually and behaviorally.

## 2026-07-13 Resizable Columns And Numbered Paging

### Problem

- Family/System list columns were fixed and could not be resized by dragging the header.
- The 150-row navigation was a small previous/range/next cluster beside the search box, so it looked temporary and did not reveal how many pages existed or allow a direct page jump.
- Moving the pager into the existing status bar required separating the status message; existing `InnerText` updates would otherwise erase nested paging controls.

### Changes

- Added IE-compatible resize handles to every fixed header cell. A drag measures the current table, applies the changed pixel width to both fixed-header and body `colgroup` elements, updates the total table width, enforces minimums, and persists by list/column-count.
- Added `goToRowWindowPage(...)`, current/total page summary, direct numbered buttons, long-page ellipsis, localized Previous/Next labels, and row range. Checked rows remain stored outside the 150-row DOM window.
- Removed paging from both search-control renderers and added it to the right side of the bottom screen-transition status bar. Status updates now target `dashboardStatusText`, preserving page controls.
- Increased the browser status reservation from 34px to 50px and added compact 1280px behavior without hiding the numbered navigation.

### Verification

- Backup: `_backups\resizable-columns-numbered-pager-20260713-095810`.
- Focused 2025 report: `artifacts\family-browser-ui-audit\20260713-100630-quality-gate`; static/contract, 2,000-row performance/cache, and all Korean/English IE scenarios passed.
- Full report: `artifacts\family-browser-ui-audit\20260713-100935-quality-gate`; all seven stages passed for 2019/2021/2023/2025/2027 with zero failures. Revit 2021 runtime is not installed and remains the expected runtime skip only.
- Stage/ProgramData SHA256: 2019/2021/2023 `49785B35FAB6EA8B474E123B48843776006C8FA6A2EDC0CE07A26624D0B9E663`; 2025 `A0467536B434EF9AFE163AD45540BBAD82CB2D143670B8EA4F7F724D389C75AD`; 2027 `4F5D4F5709687ECEF0DF34FC3A77614F835244ACDA9D7819895C8AEF7E864C6B`.
- Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_resizable-columns-numbered-pager-20260713_Setup.exe`, SHA256 `BB361BDC2FC1301225426A63CFD1BCEC3BF2E13B30AB350FFA74FBA13E1976E8`.
- Mail package: `artifacts\family-browser\mail-packages\20260713_04.zip`, 15.5 MB, SHA256 `FC79AF599902F49C2AA57FFCC846E8205E785CF962B9763686BA20E0D6B63DF6`.
- `Needs Revit Check`: with a real list over 150 rows, drag several header boundaries at workstation DPI, close/reopen the browser to confirm width restoration, then select numeric pages and confirm the bottom status/pager remains readable at the actual window size.

## 2026-07-13 Unified Family Parameter Table

### Problem

- Family detail showed the selected type parameter table separately from the family/instance/sample-type parameter table below it.
- The first type values appeared in both areas, so the same Width/Height/etc. could look duplicated and it was unclear which table was authoritative.
- The compact parameter preview was also limited to shared parameters and fixed `Take(...)` counts, so captured non-shared family/instance values could be omitted.

### Changes

- Changed `BuildParameterPreviewText(...)` in all three hosts to use all captured nonempty parameters instead of shared-only rows, and removed the family/instance/type preview row limits.
- Replaced the split main and detached-detail parameter renderers with one unified table containing family, instance, CSV, and the currently selected type rows.
- Kept one type dropdown above the table. Changing it replaces only type-scoped rows while family/instance/CSV rows remain visible.
- Deduplicated by normalized `scope + parameter name`; the sample type rows from the summary are suppressed when the structured selected-type record exists.
- Preserved lookup CSV summary rows in the unified table and assigned them an explicit `CSV` scope instead of inheriting the preceding Type scope.
- Added audit fixtures with two types (`600x600`, `1200x300`) plus family `KKY`, instance `SUP-01`, and lookup CSV data.
- Extended the IE harness to require exactly one unified table and one dropdown, then switch to the second type and verify Width `1200`, Airflow `900 CFM`, instance `SUP-01`, and CSV summary persistence.

### Defects Found By The New Gate

- First focused run found that the legacy text parser carried the Type scope into the following Lookup CSV section, so the CSV row was removed with the sample-type rows. The merge now recognizes and preserves CSV rows.
- Second focused run found that the main dashboard lacked the `paramCsvLabel` variable that already existed in the detached window. This caused a render exception and raw-text fallback. The main label is now explicitly defined and statically guarded.

### Verification

- Backup: `_backups\unified-parameter-table-20260713-102720`.
- Final focused 2025 report: `artifacts\family-browser-ui-audit\20260713-unified-parameter-focused-v3`; static/contract, build/stage, 2,000-row performance/cache, and Korean/English IE DOM/click/detail checks passed.
- Full report: `artifacts\family-browser-ui-audit\20260713-unified-parameter-full`; all seven stages passed for 2019/2021/2023/2025/2027, including ProgramData install and installed verification. Harness result: 113 scenarios, zero failures; Revit 2021 runtime remains the expected runtime skip.
- Stage/ProgramData SHA256: 2019/2021/2023 `7975B05303CF2D92373ECE505BE22BEC1ABA222D93A66B466750537D600CFC68`; 2025 `9A3C6663D1200118632B2429CD859BD5701F2D70B9E5BEBB62653EA9236F494B`; 2027 `FD49053CD84141ACF87681132CA440D4BD4F889C7B2BE9CB6C15A7E4447E8F23`.
- `Needs Revit Check`: open one real loadable family containing family, instance, multiple type parameters, formulas, and a lookup CSV. Confirm there is one table, changing the type dropdown changes only type values, and the common/instance/CSV rows remain visible without duplication.

## 2026-07-13 System Type Routing Preference Detail Rework

### Problem

- The System Type detail identity panel displayed `Segment` and `Material` from `SystemTypeSemanticSnapshot`. Those summary fields are not the routing parts registered in Revit's `RoutingPreferenceManager`, so real routing rules could exist while the identity values still rendered as `-`.
- Captured system data was split across generic text sections for routing rules, segments/sizes, dependent loadable families, layers, and a separate bottom `System Type Review Data` preview. The duplicated presentation was difficult to scan and did not resemble Revit's Routing Preferences dialog.

### Changes

- Removed the misleading identity `Segment` and `Material` rows. Basic information now keeps only valid non-empty system identity values such as class, classification, and shape.
- Added structured `@route` records sourced from each actual `RoutingPreferenceManager` rule: group, priority, part/family type, Revit class/category, registered size count, and size criterion/range. The legacy `@row` records remain in the raw snapshot only as a backward-compatible fallback.
- Integrated routing dependencies into the same routing model and fixed per-group dependency priority counting when multiple loadable parts share one role.
- Replaced the generic section renderer with one Revit-style routing-preference table. Group bands separate Segments, Elbows, Junctions, Transitions, and other rule groups; each rule has Priority, Part / Family Type, Revit Class / Category, Size Count, and Size Criteria columns.
- Removed the separately visible `Segments / Sizes` and `Dependent Loadable Families` sections. Layer information, when present, is kept inside the same unified routing surface instead of the old bottom review area.
- Hid the separate bottom `System Type Review Data` block for System Type details in both the main browser and detached detail window.
- Existing cached `@system-detail-v1` data remains readable: the renderer derives the unified table from legacy routing/dependency rows when structured `@route` records are absent. A new precise scan is needed only when an old snapshot never captured routing data at all.

### Automated Verification

- Backup: `_backups\system-routing-preference-table-20260713-111741`.
- Static guards now require the structured routing writer/parser/table, reject the old identity Segment/Material writers, and require the System Type bottom preview to stay hidden.
- IE `WebBrowser` harness selects a real audit System Type row and requires exactly one routing table, two deduplicated rule rows, group bands, segment and elbow parts, size count `12`, `min=100 / max=300` criteria, and layer content. It fails if the legacy split sections, Segment/Material dash rows, or bottom review block remain visible.
- Focused 2025 gate: `artifacts\family-browser-ui-audit\20260713-system-routing-focused-v2`; all static/contract, build/stage, 2,000-row performance/cache, and 28 Korean/English DOM/click/layout scenarios passed.
- Full gate: `artifacts\family-browser-ui-audit\20260713-system-routing-full`; all seven stages passed for 2019/2021/2023/2025/2027, including ProgramData installation and installed verification. Harness: 112 scenarios, zero failures. The 56 warnings are expected handled-alert observations from safe no-selection clicks.
- Stage/ProgramData SHA256: 2019/2021/2023 `D42190EB70C091D89DEDEF891F30D21CF88FA32C8A5B0FC8E18290191E89427C`; 2025 `12ECC360A9CBEDBFD216324752BB7E403B5FA0274675534D8A1F9D05EC214BD9`; 2027 `752CFE996AEF9E2F0DFB0280887BB3A0D055A1855C5ECF0A76C6AB0716138ED1`.
- `Needs Revit Check`: precise-scan one standard RVT containing duct/pipe/electrical system types and visually compare group order, part names, size counts, and size criteria against Revit's Routing Preferences dialog. Confirm system layer rows appear only when the selected type actually has layer data.

## 2026-07-13 Rich Result Dialog Unification

### Problem

- Many successful scan, comparison, registration, tracking, diagnostic, load, and apply operations still passed a long newline-delimited string to a result window. Even where the outer window used HTML, unstructured text remained visually flat and important counts were buried beside paths and diagnostic prose.
- The Current Model Check result still used a native multiline `TextBox` body and native export/footer controls.
- Seven ribbon-command result paths still called Revit `TaskDialog.Show` directly: standard family apply, system type apply, current model compare, system type preflight, standard RVT registration, diagnostics, and tracking stamp.

### Changes

- Added automatic result parsing to `FamilyBrowserMessageHtmlRenderer`. Ordinary `Label: Value` lines are now classified into:
  - numeric summary cards with success/info/warning/error tones,
  - project/standard/context information tables,
  - output/report path tables with monospace wrapping,
  - grouped notes and intermediate headings.
- Automatic structuring activates only when at least two result fields are detected. Short notices remain compact instead of becoming unnecessarily complex.
- Added `FamilyBrowserResultDialog` as the common result/confirmation entry point and migrated all seven ribbon-command result paths in the 2019-2023, 2025, and 2027 hosts away from direct `TaskDialog.Show`.
- Result messages containing a local or UNC output path now show `Open folder` and `Copy path` actions in the common HTML footer. `Copy details`, close, keyboard, drag, and maximize behavior remain in the same shell.
- Converted the active Current Model Check result to `FamilyBrowserHtmlDialogHost`; its title, body, footer, close button, and `Export Excel` auxiliary action are now rendered/handled through the common HTML shell. The old `CurrentModelCheckResultDialog` class remains unreferenced compatibility code and is statically guarded against reactivation.
- Preserved the specialized Family Load and System Type Apply result dialogs because they already provide item-level result tables; their older green palette was replaced with the same KKY navy/blue result styling.
- Structured error messages retain their explicit failure reason, next action, administrator information, and technical-detail sections.
- Intentional exception: `FamilyBrowserNativeCommandGuardService` continues to use Revit `TaskDialog` for immediate native-command blocking inside Revit command events. These are security/block notifications rather than operation result windows, and replacing them with a modeless browser host could weaken command interception reliability.

### Automated Verification

- Backup: `_backups\rich-result-dialogs-20260713-115447` (34 source/audit files before edits).
- Static guards require automatic result semantic markers, metric/context/output regions, report-path action resolution, the full-HTML Current Model Check route, HTML Excel auxiliary action, and all seven ribbon command migrations. Reintroducing `TaskDialog.Show` in those command files now fails the gate.
- Added Korean and English `auto-result-message` IE `WebBrowser` scenarios. They require project/standard context, three numeric cards, an output-path table, `Open folder`, `Copy path`, copy-details, close, and accept controls; visible Hangul in English or horizontal overflow fails the gate.
- Focused 2025 report: `artifacts\family-browser-ui-audit\20260713-rich-results-focused-v3`; static/contract, build/stage, performance/cache, four message/result scenarios, and 28 dashboard scenarios passed with zero failures.
- Full report: `artifacts\family-browser-ui-audit\20260713-rich-results-full`; all seven quality-gate stages passed for 2019/2021/2023/2025/2027, including ProgramData install and installed verification. Harness: 120 scenarios, zero failures; eight structured-message and eight automatic-result scenarios passed. The 56 warnings are expected handled-alert observations from safe no-selection clicks.
- Visual capture check: `artifacts\family-browser-ui-audit\20260713-rich-results-full\harness\Rvt2025-auto-result-message-ko-light\message-visual.png`; the Korean automatic-result dialog renders separate context, metric-card, report-path, note, and footer-action regions without clipping or horizontal overflow.
- Stage/ProgramData SHA256: 2019/2021/2023 `EB1E2B48572E2BFBC2AC5431A48E6A0A7857F0507545F97F27D03E37D90C5B38`; 2025 `133F0DDC1DFA1D6526D338AE0563A4842098305B746AFA957D72DD22D1976F1F`; 2027 `3D0E7D32AF0F129249B32C2115EE51BD7450BEE220B1E61676BC8B460D6601F8`.
- `Needs Revit Check`: trigger one Current Model Check with export rows, one precise-scan completion/error, one family load result, one system type apply result, and one ribbon registration/tracking result. Confirm counts are immediately scannable, long paths stay in the output section, Excel export remains available, and no active result window falls back to a flat multiline string.

## 2026-07-13 Lookup CSV Display-Case Preservation

### Problem

- Lookup CSV / Revit family size-table names were captured only through `NormalizeLookupToken`, which intentionally lowercases fingerprint content. The detail parameter table then reused that normalized signature source, so source names such as `Audit_SizeTable` appeared as `audit_sizetable`.

### Changes

- Kept the existing lowercase normalized lookup signature for fingerprint comparison, so fingerprint stability and lookup CSV difference detection are unchanged.
- Added case-preserving `lookup-display-table=` records to signature debug metadata. They store the original Revit table name plus explicit row/column counts outside the hashed signature source.
- Updated the detail preview parser to prefer case-preserving display metadata and fall back to the legacy normalized signature for older scan files.
- Removed `Normalize(...)` from the display parser path and added case-insensitive control-token matching without changing the table name itself.
- Added mixed-case `Audit_SizeTable` fixtures and assertions for the unified parameter table and Fingerprint difference table. A lowercase-only rendering now fails the UI gate.

### Verification

- Backup: `_backups\lookup-csv-display-case-20260713-123907` (12 files).
- Full report: `artifacts\family-browser-ui-audit\20260713-lookup-csv-display-case`; all seven quality-gate stages passed for 2019/2021/2023/2025/2027, including build/stage, 2,000-row performance, HTML/click simulation, ProgramData install, and installed verification.
- Harness: 121 scenarios, zero failures, 57 expected handled-alert warnings. The rendered parameter table and Fingerprint detail both retain `Audit_SizeTable` exactly.
- Existing signature files contain only the lowercase normalized name and cannot reconstruct discarded capitalization. `Needs Revit Check`: run a precise standard RVT scan to create the new display metadata; rerun the current-model precise comparison when project-side lookup CSV casing also needs to be refreshed.

### Installer

- Created: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_csv-case-20260713_Setup.exe` (3.33 MB).
- Installer SHA256: `1676ACD31EAB3851C3DEA7965344B2E0FEF851518DEB689E829DA04ABAA3CA03`.
- Mail package: `artifacts\family-browser\mail-packages\20260713_05.zip` (15.60 MB).
- Mail package SHA256: `E0E2CE65320BEF1E5CD047D7AC7A534DDE5936E2CCB6D28A2D0FDBF27E84E0C1`.
- Validation: the ZIP `Setup.exe` hash matches the standalone installer; silent installer execution returned exit code 0; installed verification passed for 2019/2021/2023/2025/2027.

## Action Inventory Snapshot

Generated from `KKY_FamilyBrowser_RevitHost_2019-2023\FamilyBrowserDashboardHtmlForm.cs` during the 2026-06-29 23:47 KST pass.

### Literal / Generated Dashboard Actions

- Global/navigation: `tab/*`, `refresh`, `lang-ko`, `lang-en`, `admin-mode-on`, `admin-mode-off`, `debug-log`
- Detail/preview: `detail-window-open`, `detail-window-sync`, `preview-inline/*`
- Standard target: `browse-discipline-*`, `discipline-*`, `mode-separated`, `mode-integrated`
- Standard RVT/admin: `open-standard-registration`, `register`, `standard-rvt-manager`, `standard-rvt-manager-close`, `standard-rvt-reset`, `standard-rvt-fast-check`, `standard-rvt-full-precise`, `standard-rvt-selected-precise`, `standard-discipline-add`, `standard-discipline-rename`, `standard-discipline-delete`
- Standard list: `standard-list-excel`, `standard-list-template`, `standard-list-export-rvt`, `standard-list-clear`, `standard-fingerprint-audit`
- Family/system apply: `sync-families`, `sync-family/*`, `sync-families-selected/*`, `apply-systems`, `sync-system/*`, `sync-systems-selected/*`
- Request workflow: `request-new`, `request-refresh`, `request-family`, `request-system`, `request-update`, `request-status/*`, `request-delete/*`, `request-attachments/*`, `request-unregistered-family/*`, `request-unregistered-system/*`, `open-requests`, `request-store-network`
- Audit/model check: `stamp`, `tracking`
- Permission/file guard/governance visible actions: `homepage-security-refresh` top pill, `file-guard-config`, `file-guard-export`
- Help / technical operations: `deployment-refresh`, `deployment-profile`, `deployment-url`, `open-managed-root`, `output`, `settings-reset-all`
- Generated export/report actions: `unregistered-families-export`, `unregistered-systems-export`
- Sub-window-only schemes: `kkyfb-request:*`, `kkyfb-select:*`, `kkysheet:*`, `kkyfileguard:*`

### Router Notes

- Main dashboard actions route through `BrowserNavigating` -> `DispatchOrRunDashboardAction` -> `RunDashboardAction`.
- UI-only actions are allowed to run directly without the modeless Revit dispatcher.
- Revit-mutating actions go through the modeless dispatcher unless no dispatcher is attached.
- Dynamic prefixes are handled before the switch: `tab/`, `sync-family/`, `sync-families-selected/`, `sync-system/`, `sync-systems-selected/`, `browse-discipline-`, `refresh-target/`, `preflight-target/`, `sync-families-target/`, `sync-systems-target/`, `discipline-`, `request-status/`, `request-delete/`, `request-attachments/`, `request-unregistered-family/`, `request-unregistered-system/`, `preview-inline/`.
- `manager-refresh` is intentionally handled inside the Standard RVT manager sub-window, not by the main dashboard router.
- Router still has hidden/legacy cases with no current visible `kkyfb:` href in the dashboard: removed manual-managed-folder actions (`policy-managed`, `policy-local`, `temporary-register`, `team-shared-test`, `request-store-local`), hidden permission/governance actions (`permission-check`, `security-*`, `permission-excel*`, `project-rule*`, `file-guard-clear`), connector request-store placeholders (`request-store-sharepoint`, `request-store-cloud`, `request-store-api`), and older maintenance aliases (`images`, `preflight`, `register-precise`, `standard-rvt-selected-images`). Treat these as compatibility/admin-internal routes unless a visible split-tab button is deliberately introduced.

## Static QA Gate

- Run before build/install: `.\KKY_FamilyBrowser_Compile\Test-FamilyBrowserUiStatic.ps1`
- Full quality gate: `.\KKY_FamilyBrowser_Compile\Invoke-FamilyBrowserQualityGate.ps1 -Years @('2019','2021','2023','2025','2027')`
- Last run: 2026-07-13 12:55 KST. Static/contract including case-preserved lookup CSV display metadata, full-HTML dialogs, live mutation refresh, bottom numbered paging, synchronized resizable columns, and the unified parameter table; five-target build/stage; staged verification; 2,000-row performance/cache/paging/resize checks; 121 Korean/English UI/message/layout/detail scenarios; ProgramData install; and installed verification passed with zero failures. Revit 2021 remains the expected `SKIP runtime-not-installed`; its build/stage/install package passed. Stage and installed DLL hashes match all five targets. The managed-data audit remains a separate external-state check because the homepage candidates were unavailable on this PC.
- Current checks:
  - `FamilyBrowserUiAudit.contract.json` verifies generated `kkyfb:` actions, exact routes, prefix routes, browser-only functions, required source tokens, and forbidden regression tokens across the 2019-2023, 2025, and 2027 hosts.
  - WinForms WebBrowser click harness renders audit scenarios without Revit, clicks visible browser-only controls, records host action candidates, checks layout overlap, and writes per-scenario JSON/HTML/stdout/stderr artifacts.
  - Structured-message harness renders the complete shared result/error dialog in Korean and English, requires full HTML title/close/body/footer/status/copy/accept elements plus cause/action/admin/technical/log/support-code regions, checks language purity and horizontal overflow, and verifies the expected visible action count.
  - Automatic-result harness feeds an ordinary newline-delimited Current Model Check result through the common renderer and requires numeric metric cards, context and output-path tables, report actions, Korean/English purity, and no horizontal overflow.
  - Performance harness records full row-store count, actual DOM row count, and visible row count separately; it fails when all rows remain hidden in the DOM, when a page exceeds 150 rows, or when page 1/page 2 checked state and Clear Selection state are inconsistent.
  - The same 2,000-row harness requires visible numeric page links, directly opens page 2, checks its active state and summary, counts header resize handles, changes a column width through the audit seam, and requires equal header/body `colgroup` pixel widths.
  - `Test-FamilyBrowserManagedData.ps1` resolves homepage candidates in runtime order and read-only validates policy, registrations, snapshots, project cache records, V2 manifests/artifacts/detail files, SHA256, and project-state references. `Invoke-FamilyBrowserQualityGate.ps1 -ManagedDataAudit` records unreachable external paths separately; add `-FailWhenManagedDataUnavailable` when a connected management folder is required by the environment.
  - Family/System detail harness resets virtual filters before inspecting a deliberately non-default fingerprint-difference row, so filtered-out rows are not incorrectly expected to remain in the DOM.
  - Setup-state harness verifies that `RVT missing` renders only the RVT registration prompt/action and `RVT connected + standard list missing` renders only the approved-list prompt/action in both Family and System panes, Korean and English.
  - Search-focus harness focuses the Family/System search box, runs queued live filtering, verifies the detail panel follows the filtered first row, and fails if search emits `detail-window-open`, `detail-window-sync`, or `preview-inline/*` host actions.
  - Detail-content harness selects an audit family row and verifies selected detail name/category, family type/composition, captured parameter values, and rendered 3D preview image markup/data URI.
  - Unified-parameter harness requires exactly one table and one type dropdown, rejects duplicate selected-type rows, switches from `600x600` to `1200x300`, and verifies that type values update while family, instance, and lookup CSV rows remain.
  - System-detail harness selects an audit system type row and verifies one Revit-style routing-preference table, grouped/deduplicated routing rules, part/class/category, size count/range, integrated dependency and layer data, and the absence of legacy split sections and bottom review content.
  - Static scan-dialog guard checks fail if precise-scan diagnostics stop writing XLSX, if OK/Cancel action fields are removed, if Delete Instance/Delete Type detection is removed, or if native dialog handling reintroduces an OK fallback without an actual OK button.
  - Project subtitle harness verifies compact local/central file-name tokens and hover `title` paths, plus Korean -> English translation of the unsaved-project placeholder.
  - English dashboard scenarios create cached display state in Korean before switching to English. Visible Hangul, cached standard/permission/tracking leaks, stored discipline-label leaks, and Korean `data-window-title` fail the gate; only the intentional `한국어` switch-button label is allowed.
  - Harness suppresses native Revit DLL popups, adds Revit language folders plus Autodesk Shared RealDWG/Components dependency directories, and force-exits so modal/native DLL issues cannot hang the quality gate.
  - Static guard rejects unescaped generated JavaScript newline regex patterns such as `replace(/\r?\n/g` inside C# string output.
  - `lang-ko` / `lang-en` routes and handlers exist in all 3 host files.
  - Deferred detached-detail language refresh field exists.
  - Shared result/message dialogs use one borderless full-HTML host and no longer contain visible WinForms title/footer/action controls or the legacy plain multiline `TextBox`; long technical detail scrolls inside its own region and structured messages expose copy-details.
  - Current Model Check review export dialog is `.xlsx` only and the export service writes only an Excel workbook.
  - Current Model Check review-result CSV writer/filter is absent.
  - File-guard export dialog does not accidentally expose unsupported CSV.
  - System type review rows populate difference export fields instead of blank difference columns.
  - System type detail rows propagate `@system-detail-v1` captured data from semantic/standard/project/comparison/preflight snapshots into main and detached detail table renderers.
  - Loadable-family precise fingerprint includes Revit family lookup CSV / size table capture through `FamilySizeTableManager`, the two-argument `GetFamilySizeTableManager(Document, ElementId)` call, and `AsValueString` cell reads.
  - Lookup table signature differences classify under `lookup tables`.
  - Host-generated HTML includes UTF-8 meta markers.
  - Previous detached-detail regressions `selectRow(this,true)` and `kkyOriginalSelectRow` are absent.
  - Detached detail window blocks both `kkyfb:` and `about:kkyfb:` navigation.
  - Dashboard shell still exposes UI-state capture and local tab switch functions.
  - Model Check target switching preserves captured `mainCenter` / active workflow scroll, and the harness performs a real capture-reset-restore round trip.
  - Admin Standard target chips follow the actual current browse target in separated mode; trade management stays beside the selector; Baseline/List action rows are checked for equal button widths/heights.
  - Queued Family/System search validation waits for the IE row window to finish rendering instead of failing on one fixed-delay sample.
  - Debug Log / F12 rendering uses `CanSeeInternalPaths()`, host toggle verifies `#fbDebug`, and JS routes missing-panel F12 to the host `debug-log` action.

## Next Work Queue

1. Revit runtime check: for one selected target, register/scan the standard RVT without connecting its standard list and confirm Family/System tabs show `등록된 표준 RVT의 표준 목록을 연결해주세요` plus the list-registration button. Switch to a genuinely unregistered target and confirm only the standard-RVT registration prompt/button appears.
2. Revit runtime check: in Model Check, scroll to several positions and switch the review target; confirm the same position remains. In Admin Standard, switch targets and confirm the current-target text and pressed chip agree, trade management is beside the selector, and both action cards remain aligned at the real DPI/window size.
3. Managed-folder/Revit runtime check: map `I:` or publish a reachable homepage candidate, run `Test-FamilyBrowserManagedData.ps1 -TreatUnavailableAsFailure`, then measure real cold/warm browser opening. Confirm the first open already uses the homepage policy, latest standard/project records are restored, the second open uses local V2 cache, and offline mode is visibly read-only.
4. Revit runtime check: register/precise-scan a standard RVT containing a model family, confirm a thumbnail PNG is created under the thumbnail cache, open Family tab detail, and confirm the detached detail window shows the same captured 3D image plus family/type parameters.
5. Revit runtime check: use a loadable `.rfa` with an imported lookup CSV / size table, change a size definition/table row/column between standard and project copies, run precise Current Model Check, and confirm the row is different with a `lookup tables` difference detail.
6. Revit runtime check: precise-scan a standard RVT with duct/pipe/electrical system types, open System Type detail, and compare the unified table directly with Revit Routing Preferences. Confirm rule-group order, priority, part/family type, Revit class/category, size count/range, and optional layer rows; confirm there are no separate Segment/Size, dependency, or bottom review sections.
7. Revit runtime check: run Current Model Check with at least one different loadable family and one different system type, export `.xlsx`, and confirm `차이 항목 / 표준 / 프로젝트 / 차이 요약` are populated in the Excel workbook.
8. Revit runtime check: switch Korean -> English -> Korean and verify the standard summary, permission/tracking pills, unsaved-project label, native window title, list/tree trade labels, result messages, and detached detail title/content all change immediately. In English, only the intentional `한국어` switch-button label may remain Hangul.
9. Revit runtime check: run safe/cancel paths for representative visible buttons only (`open-requests`, `homepage-security-refresh`, `open-managed-root`, `file-guard-config`, `file-guard-export`). Do not test hidden `permission-check` as a visible button.
10. Revit runtime check: trigger Current Model Check, precise-scan completion/error, Family/System apply confirmation/result, standard RVT registration, tracking stamp, and diagnostics. Confirm numeric cards, target information, item tables, notes, and output paths are visually separated; title, close/maximize, copy/open-folder/copy-path, Excel export, and OK/Yes/No remain usable. Verify drag/maximize, Enter, Esc, and `X`; closing a Yes/No dialog must cancel instead of approving.
11. Continue deeper trace on legacy/hidden router actions only if they become visible again or user wants them restored.
12. Update the Button Audit Table after each traced action group.
13. Revit runtime check: confirm the installed browser exposes no theme button or light/dark wording, ignores any stale `theme.txt`, and keeps the same fixed palette in the main browser, detached detail, and one structured result window at the workstation DPI.
14. Revit runtime check: open KKY Tool and Family Browser once without resizing. Confirm Family Browser starts at the same `1400 x 900` practical size, is centered on the Revit monitor, and on a constrained screen stays within 93% of the working area with no startup clipping.
15. Immediate external-PC Revit check: install the `live-state-refresh-20260713` build, rename one standard trade, and confirm the current-target label, selected chip, header standard pill, and all trade selectors change immediately without pressing Refresh. Repeat one mode switch and one standard-list connect/clear operation to confirm the same live-store refresh path.
16. Revit runtime check: open a Family or System list with more than 150 filtered rows, drag several header boundaries and confirm only the dragged column changes while every untouched header/body column remains fixed, then confirm alignment plus width restoration after closing/reopening. Use Previous/Next and direct numeric pages in the bottom status bar, confirming the current page, total pages, row range, and cross-page checked selections at the workstation DPI.
17. Revit runtime check: open a real loadable family with family, instance, two or more types, formulas, and a lookup CSV. Confirm the parameter section contains one table, the type dropdown changes only the type rows, and common/instance/CSV rows remain without duplication in both the main and detached detail views.
18. Revit runtime check: after a new precise scan, use a lookup CSV table named with mixed casing such as `Audit_SizeTable`. Confirm the unified parameter table and Fingerprint difference table preserve that exact spelling while content comparison still detects row, column, and value differences.
19. Revit runtime check: precise-scan a loadable family containing both an ordinary Integer parameter and a Yes/No parameter. Confirm ordinary values remain numeric while the Yes/No value shows `Yes` or `No` in the main and detached detail parameter tables, then switch family types and confirm the value follows the selected type.
20. Revit runtime check: open a real pipe or duct system type whose Routing Preferences include bounded and unbounded size criteria. Confirm the detail range is a `조건 / 최소 / 최대` table, defaults to `mm`, switches immediately to `in`, and shows Revit sentinel bounds as `제한 없음` instead of scientific notation. In Korean, confirm the selected groups read `엘보 (Elbows)`, `접합 (Junctions)`, `전이 (Transitions)`, `유니온 (Unions)`, and `캡 (Caps)`; in English, confirm only the English group names appear.
21. Revit runtime check: apply one system type absent from the current project and load one absent standard family. Confirm the confirmation/result item tables label the operations `신규 생성` and `신규 로드` (English: `Create new`, `Load new`) while genuinely skipped or unavailable items still use the semantic `누락` wording.
22. Revit runtime check: load one absent Family and apply one absent System Type. Confirm each row immediately becomes `로드됨 · 저장/동기화 대기` / `적용됨 · 저장/동기화 대기`, is not selectable, and is excluded from `Load Available`. First close the project without saving and reopen it to confirm the persisted list returns to `Load Available`. Then repeat and test successful Save, Save As, and Synchronize with Central; each must remove the temporary state and rebuild the row from the persisted project. A cancelled/failed save and a cancelled close must leave the temporary state intact until a later successful commit or completed unsaved close.
23. Revit runtime check: on a PC with the homepage-managed folder connected, add the central RVT to File Guard with both family and type blocking enabled. Switch Admin Mode OFF and confirm the switch completes without a full list reload, Revit Load Family is cancelled, and family/type rename is blocked for both the central file and a differently named local copy. Switch Admin Mode ON and confirm both operations are available again. On a disconnected PC, confirm the settings action explains that the managed shared folder is unavailable and gives the local diagnostic log path instead of `(managed log unavailable)`.
24. Revit/runtime network check: open Family Browser on a first-use PC where every homepage-managed candidate is unavailable. Confirm Home shows the structured missing-network-folder notice once, a general user can close and contact the manager/distributor, and a management-only user can select an existing writable UNC or mapped network share. Confirm local folders are rejected, the generated root README warns against manual edits, the next browser open reuses the saved managed root, and long clipped labels/path cells reveal their full text on hover while widened list columns reveal as much inline text as fits.
25. Revit runtime check: run Standard RVT precise scan and Current Model Check with a family that raises `Opening not cutting anything` while `EditFamily` opens it. Confirm the enabled OK button is clicked automatically, the scan continues, and the result lists the category/family/action/reason. Also test a disposable family whose dialog exposes only Delete Instance/Delete Type/Delete plus Cancel; confirm Cancel is chosen, no destructive button is clicked, and the on-demand Excel export contains the `스캔경고` / `ScanDialogs` sheet with the full dialog text. Revit 2021 remains build/Stage-only until its runtime is installed.
26. Revit runtime check: prepare at least two separated trades (for example Architecture and Mechanical), each with a registered/scanned standard RVT and connected approved list. Leave Mechanical selected in Standards, open both Family and System Type lists, then click the Architecture ready chip below search. Confirm the chip changes immediately, Mechanical rows disappear during loading, Architecture rows replace them without pressing Refresh, and the saved Architecture Current Model Check state/differences are restored. Switch both ways, then rescan or reconnect one target and confirm the next switch uses the new data rather than a stale prepared/row cache.
27. Revit runtime check: precise-scan a disposable standard RVT containing at least one wall, floor, and roof type with multiple compound layers, a core region, a structural material layer, and a variable layer. Open System Type detail and compare row order, exterior/interior direction, function, material, thickness, core boundaries, and badges directly with Revit Edit Assembly. Confirm first use shows `mm`, switching either routing or layer control to `in` updates both sections, closing/reopening restores `in`, switching back to `mm` persists, and the main/detached detail windows remain synchronized. Also open one pre-change snapshot without rescanning and confirm its legacy layer strings still render as a usable table.
28. Revit runtime check: precise-scan a disposable standard RVT where parent A contains child a/b/c, child c has no project-level standalone instance, and another nested family d also has one standalone instance. Register the target project in File Guard and enable `하위 전용 패밀리 단독 모델링 금지`. With Admin OFF, confirm A placement succeeds and generated a/b/c remain, direct c placement and copy/paste are rolled back, and direct d placement remains allowed. With Admin ON or the option disabled, confirm direct c placement is allowed. Repeat with a differently named local copy of the guarded central project.
29. Revit runtime check: create curtain-wall and curtain-system standards using both a System Panel and a loadable Curtain Panel as the default panel. Run a new precise standard scan and Current Model Check, then change the default panel, one panel type parameter, and one referenced/dependent family type in the project copy. Confirm the parent Curtain Wall/Curtain System becomes `표준과 다름`, the detached System detail shows `커튼패널 의존 구성` and `커튼패널 구성 차이`, and the same comparison remains active when `System Type 상세 구성 요소 비교 (Railing, Stair 등)` is disabled. Revit 2021 remains build/Stage-only until its runtime is installed.
30. Revit runtime check: after a new schema-9 precise scan, open Railing, Stair, Curtain Wall/Curtain System, and direct PanelType details containing length-valued dependent type parameters. Confirm the detailed component and component-difference tables default to `mm`, switch every synchronized selector to `in`, preserve the selected unit after closing/reopening, and show the same converted values as Revit. Confirm disabling `System Type 상세 구성 요소 비교 (Railing, Stair 등)` hides only Railing/Stair component tables while mandatory Curtain Panel tables and their unit control remain visible.
31. Revit runtime check: enable `System Type 상세 구성 요소 비교`, run Current Model Check, return to Standards and System Types, and confirm the checkbox remains enabled and a real Railing row still shows Run/Landing/Support or rail/Baluster component tables. Restart Revit and confirm the same state. Then disable it, run Current Model Check, and confirm it remains disabled while only optional Railing/Stair sections are hidden. For a comparison report created by the older destructive-filter build, confirm selecting a Railing row restores standard V2 detail without a new scan when the V2 detail record exists; otherwise run one new precise standard scan.
32. Revit runtime check: register and accept-scan a disposable standard RVT, then modify one family/type and save, fail/cancel one save, synchronize once, and replace the RVT externally. Confirm Home/header/Standards detect source revision changes without manual Fast Stamp Check; stale Family/System rows and mutation actions remain blocked until rescan; only successful Save/Save As/Sync creates immutable item history; failed/cancelled saves create none; external replacement shows changed/unknown-author rather than inventing item history. Repeat through mapped-drive and UNC aliases and confirm they identify the same source.
33. Needs Revit Check: the per-project lightweight catalog baseline is implemented and deployed. It captures loadable family names, loadable family type names, and supported System Type names without `EditFamily`, fingerprints, parameters, CSV, images, or model mutation; observes on browser startup/activation, Refresh, and successful Save/Save As/Sync; keeps accepted and last-observed catalogs separate; shows persistent added/removed warnings; and classifies committed Browser operations versus external/untracked changes. Run a disposable-project matrix covering first baseline, browser load then Save/Sync, external Family/type add/delete/rename, failed/cancelled save, local/central aliases, and a genuinely large model. Record real API-thread capture time and confirm automatic startup observation remains acceptably responsive before treating the sub-second target as production-proven.
34. Needs Revit Check: execute the complete disposable lifecycle matrix on two PCs against one reachable managed share: first baseline, Browser Family load/System apply, Save/Save As/Sync success, cancelled and failed save, offline write followed by reconnect, external Family/type add/delete/rename, standard RVT external replacement, mapped-drive/UNC aliases, TEST-root to homepage-root migration, request create/edit/attach/state/delete, pending operation/history flush, and accepted project-catalog update. Confirm each committed mutation appears once, failed/cancelled work creates no committed history, offline pending records become visible and flush exactly once, and stale standard/project data blocks mutation until reviewed.
35. Implemented; Needs Revit/network check: request status changes and deletes now carry `Revision` plus `RevisionToken`, acquire a request-scoped same-folder lock, re-read the authoritative request under that lock, and reject stale screens instead of overwriting a newer administrator change. The browser reloads the latest request list after a conflict. Validate the real two-PC behavior through one reachable SMB share before production sign-off; see item 37.
36. Completed deployment: the request-concurrency payload is built and staged for Revit 2019/2021/2023/2025/2027, installed to ProgramData, and verified by exact DLL and `.addin` hashes. Revit 2021 remains build/install verification only because its runtime is not installed on this PC.
37. Needs Revit/network check: open the same request from two PCs against one reachable managed SMB share, change its status on PC A, then attempt stale status change and stale delete on PC B. Confirm PC B receives the conflict result, reloads the latest request, and cannot overwrite or delete PC A's change. Repeat through mapped-drive and UNC aliases. All participating clients must run this revision; an old client that does not honor the request lock cannot provide full simultaneous-write safety. A cloud-synced/local replica is not equivalent to a shared SMB lock and remains unsupported for authoritative concurrent editing.
38. Implemented: request attachments now use content-addressed same-folder temporary copies, retry without duplicate files or metadata, and roll back new attachment state when the authoritative request JSON cannot commit. Request deletion writes an immutable full `DeletePrepared` snapshot under `RequestAudit/Deleted` before removing active data and appends `DeleteCompleted` or cleanup-failure evidence. See the implementation record at the end of this file.
39. Needs network check: on one real managed SMB request store, create a request with two large attachments, interrupt or deny the first authoritative request write, reconnect, and retry. Confirm there is one stored copy per unique content hash, no stale attachment metadata from the failed store, and no `.kky-t-*` files. Delete the request and confirm active JSON/mail/attachments are removed while immutable `DeletePrepared` plus `DeleteCompleted` records retain the full request/history/attachment metadata. Repeat a cleanup-failure retry with one attachment temporarily locked.
40. Automated deep audit completed; Needs Revit/two-PC check: enable `프로젝트 요소 생성·수정·삭제 추적`, open one standalone project and one workshared local/central pair, then exercise create, modify, delete, create-then-delete, matched and grouped Undo/Redo, local Save, Save As, successful/failed/cancelled Save or Sync, cancelled close, actual close/reopen, and tracking OFF/ON. On Revit 2023/2025/2027 explicitly run Reload Latest and confirm the native bridge rebases without attributing incoming elements to the current workstation; on 2019 confirm the transaction-name fallback does the same. Force or simulate one baseline/rebase capture failure and confirm the UI retains `외부 업데이트 범위 누락 / 검토 필요` instead of later clearing the uncertainty. Repeat from two add-in-equipped PCs through mapped-drive and UNC aliases, including offline/reconnect replay, a valid TEST-root to homepage-root checkpoint migration, an invalid/corrupt checkpoint that must remain quarantined, and first-baseline time on a large model. Open `최근 변경 이력 보기`, verify local-save protection and central-publication times remain distinct, confirm no workbook is created merely by viewing it, and export XLSX explicitly. Confirm an edit made on a workstation without this add-in is reported only as an uninstrumented coverage gap/unknown external change, never as a fabricated user-level event.
41. Needs Design: define production retention, archive/compression, and date/project filtering for the append-only element history before long-term rollout. The current viewer intentionally reads the latest 200 commits and caps explicit XLSX preparation at 5,000 change rows. Also define how server-clock skew, global ordering across PCs, and simultaneous edits to the same element should be presented; the existing ledger preserves each client commit but does not claim a conflict-free global forensic order.
42. Needs Security/Operations: define the production audit threat model before treating element or policy history as forensic evidence. Decide whether to use a keyed HMAC/digital signature, an append-only server store, daily signed manifests/high-water anchors, deletion detection, restricted service-account writes, immutable backups, retention, and restore drills. Current SHA-256 projections detect accidental corruption and semantic contradiction but a user with managed-share write/delete permission can recompute plausible hashes or remove complete valid files. Operation and standard-candidate history also need a versioned integrity migration if record-level corruption detection is required.
43. Needs Revit/two-PC check: configure at least two ready standard trades, register separate test RVTs in Permissions / Guard, assign each RVT a trade directly and once through XLSX import, then reopen each project. Confirm the Model Check pane immediately shows the assigned automatic target, the first stale/missing state runs one comparison against only that trade, the second unchanged open reuses the saved result, changing the project or assigned standard triggers exactly one new check, and no XLSX is created until the user explicitly exports. Open the same guarded central project from two add-in-equipped PCs at nearly the same time and confirm one session scans while the other waits and reuses the committed cache. For a nested-only child name present in another trade, confirm direct-placement enforcement uses only the RVT's assigned trade. Revit 2021 remains build/Stage/install verification only until its runtime is installed.
44. Needs Revit performance check: with diagnostics enabled, load one and four Families, then apply one and four System Types from both a local and a real managed-network Standard RVT. Confirm the owned Standard RVT closes, the list refresh completes, and only then the result window appears; clicking OK must return immediately with no second loading cursor. In one guarded workshared project, measure a small and a large Synchronize with Central before/after this revision, confirm the Revit completion callback returns after local durable spool creation instead of waiting for managed-share history publication, and confirm the shared history appears shortly afterward or flushes exactly once after reconnect. Review `dashboard-runtime-YYYYMMDD.log` and `element-tracking-performance.log` for the per-stage timings.

## 2026-07-13 Yes/No Parameter Display

### Problem

- Family/type parameter capture used `FamilyType.AsInteger(...)` or `Parameter.AsInteger()` and persisted the raw number, so Revit Yes/No parameters appeared as values such as `1` or `2` in the unified parameter table.
- A UI-only `1 -> Yes` conversion would corrupt ordinary Integer parameters, so the fix had to identify the Revit parameter definition before formatting.

### Fix

- Backup: `_backups\yes-no-parameter-display-20260713-134402` (18 existing source/audit files).
- Added shared `KKY_FamilyBrowser_SharedUi\FamilyBrowserYesNoParameterFormatter.cs`.
  - Revit 2019/2021 compatibility: reflect `Definition.ParameterType` and require `YesNo`.
  - Revit 2023/2025/2027 compatibility: reflect `Definition.GetDataType()` and require a boolean/yes-no ForgeTypeId.
  - Mapping is `0 -> No`, nonzero -> `Yes`; a non-Yes/No Integer keeps its existing formatted or numeric value.
- Routed every persisted display-value capture through the same formatter in all three host source trees:
  - `FamilyDocumentParameterCaptureService`
  - `StandardLibraryRegistrationService` family-document and standard-RVT element capture
  - `ProjectSnapshotCaptureService`
  - `SystemTypeSemanticCaptureService`
- Added an audit family fixture whose first type contains `IsEnabled = Yes` and second type contains `IsEnabled = No`; the unified parameter harness now verifies both values across the type dropdown change while still checking numeric Width values.
- During the full gate, the 2023/1280/English search-focus test exposed an unrelated intermittent row-store timing race. The harness now derives its search term from the stable `KKYFB._stores` data after synchronizing current filters instead of racing a replaceable rendered `<tr>`.

### Verification

- Static and action contract checks: PASS for all three host trees (`77` generated actions, `247` exact routes, `59` prefix routes, `11` browser functions each).
- Build: PASS, zero compile errors for the 2019/2021/2023 net48 host, 2025 host, and 2027 host.
- Stage verification: PASS for 2019/2021/2023/2025/2027.
- 2,000-row performance/cache gate: PASS.
- Targeted intermittent-search reproduction after harness correction: `6/6 PASS`.
- Full HTML/IE WebBrowser gate: `120 PASS + Revit 2021 runtime-not-installed SKIP`, zero UI failures. Report: `artifacts\family-browser-ui-audit\20260713-yes-no-parameter-display-final\harness`.
- ProgramData install: BLOCKED after all code gates passed. The current non-elevated process has read-only access to `C:\ProgramData\Autodesk\Revit\Addins\2019\KKY_FamilyBrowser\KKY_FamilyBrowser_RevitHost.dll`; the install script stopped on `Access denied` before replacing installed files. Stage contains the verified new build; ProgramData still contains the prior installed build.
- Existing snapshots do not carry enough type metadata to safely reinterpret arbitrary stored integers. Re-run the relevant standard precise scan/project comparison scan after deployment to regenerate `Yes/No` display values.
- Needs Revit Check: after elevated deployment and a new precise scan, verify one real Yes/No parameter in the main and detached detail views for both false and true family types.

## 2026-07-13 System Routing Size Range Units

### Problem

- Revit routing-preference criteria are captured in Revit's internal length unit, feet. The detail view rendered that raw value directly, so a 1/2-inch minimum appeared as `0.041666...` and a 24-inch maximum appeared as `2`.
- Revit also exposes unbounded criteria with very large sentinel values such as `-1E+30` and `1E+30`; these were displayed as if they were real sizes.
- Multiple criteria were stored in one text string (`size 1 min=... max=...; size 2 ...`), which made the range difficult to scan.

### Fix

- Backup: `_backups\system-routing-unit-table-20260713-142956` (9 existing source/audit files).
- Added shared `KKY_FamilyBrowser_SharedUi\FamilyBrowserSystemRoutingUnitUi.cs` and connected it to both the main detail panel and detached detail window in all three host trees.
- Kept captured/fingerprint source values unchanged in internal feet. The display layer now parses each criterion into a compact `Criterion / Minimum / Maximum` table.
- Added a `Size unit` / `사이즈 단위` selector with `mm` as the default and `in` as the alternate. Switching units updates only the visible size cells and does not rebuild the dashboard, change the selected item, or reset scroll.
- Conversion uses `feet x 304.8` for millimetres and `feet x 12` for inches. Values are rounded for display only; comparison data remains untouched.
- Non-finite or absolute values at/above `1E+20` render as `No limit` / `제한 없음`, so Revit's `E+30` sentinels no longer leak into visible text.
- Existing snapshots that already contain `@system-detail-v1` routing criteria use the new table immediately; no rescan is required for this UI change. A rescan is needed only when an older snapshot has no captured routing-detail data at all.

### Automated Verification

- Static and action contract checks: PASS for all three host trees (`77` generated actions, `247` exact routes, `59` prefix routes, `11` browser functions each).
- Audit fixture now uses real internal-feet values: `0.3280839895013123 ft -> 100 mm -> 3.937 in`, `0.984251968503937 ft -> 300 mm -> 11.811 in`, plus `+/-1E+30` unbounded sentinels.
- The IE `WebBrowser` harness requires one unit selector, default `mm`, one nested criteria table with two rows, correct mm/in conversion, localized `No limit`, and no visible `min=`, `max=`, or `E+30` text.
- Focused 2025 report: `artifacts\family-browser-ui-audit\20260713-system-routing-units-focused`; 30 Korean/English UI scenarios passed with zero failures.
- Full report: `artifacts\family-browser-ui-audit\20260713-system-routing-units-full`; all five quality-gate stages passed, including five-target build/stage verification and the 2,000-row performance/cache gate. HTML/IE harness: `120 PASS`, zero failures, 56 expected handled-alert warnings; Revit 2021 runtime is the expected `SKIP runtime-not-installed` while its build/stage package passed.
- Stage SHA256: 2019/2021/2023 `67D022A45873E1C44824B30B9E706CFEC63357FE44D05EBEF2C7CFF44F23A38A`; 2025 `26EBDB8CB127E11D28CC10553AFF3E4058E9F03D7D717806088D3057B2EE5785`; 2027 `680DA669CD4B57DDBA53A7D0B32AFFE07C78F62176DCE8F9585516C408E36B90`.
- This run built and staged the add-ins but did not replace the ProgramData installation (`Install: False`).
- Needs Revit Check: compare one real pipe/duct Routing Preferences dialog with the Family Browser detail table, including one bounded range and one all-sizes/unbounded rule, then switch `mm` and `in` in both the main and detached detail views.

## 2026-07-13 New Load / Create Work-Item Terminology

### Problem

- A system type absent from the current project was shown in the apply confirmation/result work-item UI as `누락 생성` (`Create missing`). The item is not a failed or omitted operation; it is a valid new type creation, so the label was semantically misleading.
- No literal `누락 로드` label existed in the active family-browser source. However, an absent standard family receives `ExecutionMode = Load`, and the HTML result table displayed only `로드` (`Load`), which did not distinguish a new family load from update/reload work.

### Fix

- Backup: `_backups\new-load-create-terminology-20260713` (audit, static test, and the affected files in all three host source trees).
- Updated all three `FamilyBrowserOperationHtmlDialog` copies:
  - system-type creation summary/result: `누락 생성` -> `신규 생성` (`Create missing` -> `Create new`)
  - absent-family result action: `로드` -> `신규 로드` (`Load` -> `Load new`)
- Updated all three dashboard fallback/translation paths so the system apply summary uses `신규 지원 타입 생성` and the `CreateMissingType` display translation uses `신규 생성`.
- Kept internal routing and persisted action values such as `CreateMissingType`, `createmissingtype`, and `LoadMissingDependencyFamily` unchanged. Descriptions for genuinely missing, skipped, or unavailable data still use `누락` because that meaning remains correct.
- Extended `Test-FamilyBrowserUiStatic.ps1` to require the new Korean/English labels and reject reintroduction of `누락 생성` / `누락 로드` work-item labels in every host.

### Verification

- Static and action contract checks: PASS for all three host trees (`77` generated actions, `247` exact routes, `59` prefix routes, `11` browser functions each).
- Build and Stage verification: PASS for 2019/2021/2023/2025/2027.
- 2,000-row performance/cache gate: PASS.
- Full HTML/IE WebBrowser harness: `120 PASS`, zero UI failures; Revit 2021 runtime is the expected `SKIP runtime-not-installed` while its build/stage package passed.
- Full report: `artifacts\family-browser-ui-audit\20260713-new-load-create-terminology\quality-gate-summary.md`.
- This run built and staged the add-ins but did not replace the ProgramData installation (`Install: False`). No rescan is required because only user-facing terminology changed.
- Needs Revit Check: inspect one real system-type apply result and one new-family load result after the next deployment to confirm the two labels at workstation DPI.

## 2026-07-13 Provisional Load / Apply Commit Lifecycle

### Problem

- After a successful Family load or System Type apply, the browser could still show the row as `Load Available`, allowing the visible list state to contradict the live Revit document.
- The operation log already used `PendingSaveOrSync`, but the list never consumed it and only the existing central-sync path participated in commit handling. A successful in-memory mutation, a persisted save/sync, and an unsaved close were therefore not represented as separate states.

### Fix

- Backup: `_backups\pending-save-sync-lifecycle-20260713-153915` (23 pre-edit files).
- Added a per-open-`Document` pending operation queue in all three host trees. Successful operation outcomes are verified against the live Revit document before they are exposed as pending.
- Added source-scoped row overlays:
  - Family: `FamilyPendingSaveOrSync` -> `로드됨 · 저장/동기화 대기`
  - System Type: `SystemPendingSaveOrSync` -> `적용됨 · 저장/동기화 대기`
- Pending rows are visible but disabled, cannot be checked/applied again, and are excluded from `Load Available`. Their detail view explains that Save or Synchronize with Central is required for confirmation.
- Subscribed all host versions to `DocumentSaved`, `DocumentSavedAs`, `DocumentSynchronizedWithCentral`, `DocumentClosing`, and `DocumentClosed`.
  - Only a `Succeeded` Save, Save As, or Synchronize event finalizes the pending batch.
  - Before commit logging, the loaded Family/System Type is checked again in the live document. Missing items are recorded as `NotCommitted`.
  - Failed or cancelled commit events leave the batch pending.
  - `DocumentClosing` only remembers the pending key. State is discarded after `DocumentClosed`, so cancelling the close does not lose the temporary state.
- The pending key follows the live Revit `Document` instance instead of its path, so Save As cannot orphan the batch when `PathName` changes.
- A successful commit invalidates project row/scan caches and refreshes the open dashboard. Precise project scan publication is deferred while pending entries exist, preventing unsaved in-memory content from being written as a confirmed shared comparison result.
- Added Korean/English audit scenarios for Family and System pending rows, pending-state selection guards, detail warning checks, event/static contract guards, and a language-transition regression check.

### Verification

- Static and action contract checks: PASS for all three host trees (`77` generated actions, `249` exact routes, `59` prefix routes, `11` browser functions each).
- Build and Stage verification: PASS for 2019/2021/2023/2025/2027, zero compile errors.
- Targeted pending-state harness: `32/32 PASS`, including Family/System and Korean/English on Revit 2019/2023/2025/2027 hosts.
- 2,000-row performance/cache gate: PASS with zero failures.
- Full quality gate: all five stages PASS. HTML/IE harness: `136 OK`, one expected Revit 2021 `SKIP runtime-not-installed`, zero failures.
- Full report: `artifacts\family-browser-ui-audit\20260713-pending-save-sync-final\quality-gate-summary.md`.
- This run built and staged all five versions but did not replace ProgramData (`Install: False`).
- Needs Revit Check: execute queue item 22 in a real project. Automated Revit-free tests verify the state contract and event wiring, but cannot prove Autodesk's runtime event order for successful/cancelled Save, Save As, Sync, and close dialogs on a real workshared document.

## 2026-07-13 Korean Routing Group Labels with English Terms

### Problem

- System Type detail translated Revit Routing Preferences groups to Korean-only labels such as `엘보`, `접합`, and `전이`.
- Users working in English Revit could not quickly map those less-familiar Korean terms back to the labels used in Revit's native dialog.

### Fix

- Backup: `_backups\routing-bilingual-labels-20260713-164641` (audit, static/harness tests, and all potentially affected host/scenario files).
- Updated the main detail panel and detached detail window in all three host trees so Korean mode displays only the requested five bilingual labels:
  - `엘보 (Elbows)`
  - `접합 (Junctions)`
  - `전이 (Transitions)`
  - `유니온 (Unions)`
  - `캡 (Caps)`
- `Segments`/`세그먼트`, `Crosses`/`십자 연결`, and every other term remain unchanged. English mode continues to display the original English-only group names.
- The change is display-only. Captured `@route` group keys, rule ordering, snapshot/fingerprint data, and comparison semantics remain unchanged, so no rescan is required.
- Added static guards requiring all five bilingual mappings in both main and detached renderers for every host. The IE harness now verifies `엘보 (Elbows)` in Korean and rejects Korean leakage from the English system-detail table.

### Verification

- Static/action contract checks: PASS for all three host trees (`77` generated actions, `249` exact routes, `59` prefix routes, `11` browser functions each).
- Release build and staged add-in verification: PASS for 2019/2021/2023/2025/2027, zero compile errors. Existing 2025/2027 WindowsBase reference warnings remain warnings only.
- 2,000-row performance/cache gate: PASS.
- HTML/IE WebBrowser harness: `136 OK`, zero failures; Revit 2021 runtime is the expected `SKIP runtime-not-installed` while its build/stage package passed.
- Full report: `artifacts\family-browser-ui-audit\20260713-routing-bilingual-labels-final\quality-gate-summary.md`.
- ProgramData installation was not requested and was not run (`Install: False`).
- Needs Revit Check: visually compare one real duct/pipe Routing Preferences dialog with both the main and detached System Type detail tables in Korean and English.

## 2026-07-14 Five-Version Installer Package

### Build And Package

- Audit backup: `_backups\installer-package-20260714-0824\FAMILY_BROWSER_BUTTON_AUDIT.md`.
- Rebuilt Release hosts and regenerated the stage for Revit 2019/2021/2023/2025/2027. All five staged add-in manifests and payloads passed `Verify-FamilyBrowserRecovered.ps1`.
- Existing installer artifacts were preserved. The packaging script's old-artifact cleanup was intentionally bypassed, while using the same verified stage and Inno Setup definition.
- Inno Setup 6.7.3 compiled all five version tasks and payloads successfully:
  - `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260714_Setup.exe`
  - Size: `3,502,232 bytes`
  - SHA256: `3DBE80CF70B563A3B227DACBDDAFBDCBB5FAD03078D55D2FE13B5D6AC2A259F5`
- Created a mail-sized package above 15MB:
  - `artifacts\family-browser\mail-packages\20260714_01.zip`
  - Size: `16,013,684 bytes` (`15.27 MB`)
  - SHA256: `6D862CA5EB1AE3FCEED02D34A385F5C6F8AB5B3A7E1753CFB89092FEF156FBFA`
- ZIP content validation passed: `Setup.exe`, `README.txt`, and the mail-size padding file are present. The SHA256 of `Setup.exe` read directly from the ZIP exactly matches the standalone installer.
- ProgramData installation was not requested and was not run.

## 2026-07-14 On-Demand Result Excel Export

### Problem

- Standard precise scan and selected-family refresh wrote standard-scan-dialog-diagnostics-*.xlsx directly into the managed snapshot folder without asking the user for a destination.
- Standard/project 3D image generation also wrote a timestamped thumbnail-diagnostics-*.txt after every batch.
- Precise standard scans opened a follow-up Fingerprint Excel save prompt before the result dialog, while Family Load and System Type Apply result dialogs had no direct Excel export action.
- Repeated scans could therefore accumulate user-facing result files even when nobody needed an export.

### Fix

- Backup: _backups\on-demand-result-excel-20260714.
- Removed automatic diagnostic workbook creation from all three StandardLibraryRegistrationService copies. Auto-handled Revit dialog records remain in the in-memory result object and the compatibility DiagnosticReportPath stays empty.
- Removed automatic thumbnail diagnostic text creation from all three FamilyThumbnailPreviewService copies.
- Removed the automatic post-scan Fingerprint Excel prompt. The existing explicit Fingerprint audit command in Standards remains available and still opens a destination picker.
- Added shared FamilyBrowserResultExcelExportUi:
  - always opens SaveFileDialog
  - creates no file when the dialog is cancelled
  - writes .xlsx only after the user confirms a destination
  - reports success/failure in the HTML result footer
- Added Excel 내보내기 / Export Excel to these HTML result dialogs:
  - standard RVT fast/precise scan
  - selected standard family precise refresh
  - selected/all 3D image update
  - Family Load
  - System Type Apply
- Exported rows contain the result-specific summary and records: scan counts and auto-handled warnings, missing/updated family rows, thumbnail item outcomes, Family Load items, and System Type Apply items.
- Current Model Check and existing explicit Standard List/Unregistered/File Guard exports were already destination-driven and remain unchanged.
- Operational JSON, snapshots, preview images, standard lists, and save/sync tracking records are still written because the browser needs them to function; only optional user-facing result/diagnostic exports were changed.
- Existing diagnostic files were not deleted automatically.

### Verification

- Static/action contract checks: PASS for all three host trees (77 generated actions, 249 exact routes, 59 prefix routes, 11 browser functions each).
- Added regression guards that reject standard-scan-dialog-diagnostics-, WriteRegistrationDialogDiagnosticReport, WriteBatchDiagnosticReport(result), diagnostic path text, and the old thumbnail follow-up prompt.
- Release build and staged add-in verification: PASS for Revit 2019/2021/2023/2025/2027.
- 2,000-row performance/cache gate: PASS.
- HTML/IE WebBrowser harness: 136 scenario results, 0 failures, 72 expected handled-alert warnings.
- Full report: artifacts\family-browser-ui-audit\20260714-on-demand-result-excel-final\quality-gate-summary.md.
- ProgramData installation and installer packaging were not requested and were not run (Install: False).
- Needs Revit Check: run one precise standard scan and one 3D image update, verify no new diagnostic Excel/TXT appears before clicking Export Excel, cancel the destination dialog once, then export and inspect the workbook. Repeat once for Family Load and System Type Apply results.

## 2026-07-14 Nested Loadable Family Difference Propagation

### Finding

- The current-project precise scan already enumerates every browser-loadable `Family` in the open project, not only top-level families. Each editable family is opened with `EditFamily` and receives its own content fingerprint, parameter capture, type data, and signature debug record.
- The standard precise scan already performs the same deep capture and records each parent family's nested loadable-family relationships.
- During comparison, current-project parent/child relationships are reconstructed from the project signature records as `ProjectNestedLoadableFamilies` and standard relationships come from `NestedLoadableFamilies`.
- The missing behavior was downstream: matching and differing nested helper rows were both hidden by the list, and a nested child's independent fingerprint difference was not propagated to every parent family that used it.

### Fix

- Backups:
  - `_backups\nested-family-difference-propagation-20260714-095922`
  - `_backups\nested-family-difference-validation-20260714-continued`
  - `_backups\nested-family-ui-correction-20260714`
  - `_backups\nested-family-performance-fixture-20260714-1045`
- Added shared `NestedLoadableFamilyDifferencePropagationService` and invoked it after loadable-family comparison in all three host trees.
- Built a parent map from both standard and current-project nested relationships. A differing, missing, project-only, or manual-review nested child is now linked to every parent that uses it.
- Differing nested children are exposed in the Family list even though ordinary matching nested helpers remain hidden. The exception is allowed only when an approved displayed parent is present, so unrelated helper noise is not added to the list.
- Nested helper rows cannot be checked or loaded independently. Their action is `상위 패밀리 검토` / `Review parent family` and their detail view names the parent families.
- Every affected parent is marked different. Multi-level nesting propagates recursively to the top-level parent, and a parent that was otherwise `LoadedLatest` becomes `DifferentFromStandard` or `ManualReview`.
- The child's exact fingerprint details are copied to the parent detail model while preserving their original area. Parameter, formula, type, CSV/lookup-table, geometry, and other existing difference rows therefore keep the same table renderer instead of being flattened into text.
- Added guards for matching children, missing/project-only children, wholly absent parents, already-different intermediate parents, and recursive parent chains. A wholly absent parent does not create a redundant standalone child warning.
- Fixed four UI regressions found by the first IE harness run: nested action labels being overwritten, affected top-level parents hidden in modeler mode, propagated parameter details routed to the wrong table, and new Korean nested-family text leaking into English mode.
- Adjusted the synthetic performance fixture to count list-visible rows. The intentionally hidden matching nested child no longer turns a requested 1,000-row test into 999 visible rows.

### Verification

- Nested propagation behavior test: PASS for direct, matching, transitive, project-only, missing, absent-parent, and pre-different-parent cases.
- Static/action contract checks: PASS for all three host trees (`77` generated actions, `251` exact routes, `59` prefix routes, `11` browser functions each).
- Focused Revit 2019 IE `WebBrowser` regression run: `34/34 PASS`, including Korean/English, admin/modeler, 1280px viewport, detailed difference tables, disabled child selection, and matching-child suppression.
- Full Release build and staged add-in verification: PASS for Revit 2019/2021/2023/2025/2027, zero compile errors.
- 2,000-row gate: PASS. Family/System lists show exactly 1,000 logical rows with 150 DOM rows; filter response was `10-12ms` and warm usable time was `268-369ms` on installed runtime targets.
- Full HTML/click/language/layout harness: `136 OK`, zero failures. Revit 2021 UI runtime is the expected `SKIP runtime-not-installed`; its build and stage verification passed.
- Final report: `artifacts\family-browser-ui-audit\20260714-nested-family-difference-final-v2\quality-gate-summary.md`.
- ProgramData installation and installer packaging were not requested and were not run (`Install: False`).

### Needs Revit Check

- Use a fixture with a top-level composite family and at least one nested loadable family. Change one nested child parameter, type, lookup CSV cell, or geometry only in the project copy, then run the standard precise scan and current-model precise check.
- Confirm that the differing child and every parent in the chain appear as different, the child and parent details show the same exact difference, a matching nested child remains hidden, and the nested child cannot be loaded directly.
- The schema change itself does not force a rescan when old snapshots still contain valid nested metadata and signature debug files. For a reliable first verification after deployment, rerun the standard precise scan and current-model precise check once so both relationship sources are definitely present.

## 2026-07-14 Routing Preference Minimum Size Display

### Problem

- In System Type detail, Routing Preferences maximum sizes were displayed correctly but every minimum size appeared as `-`.
- Revit API inspection confirmed `PrimarySizeCriterion.MinimumSize` and `MaximumSize` are both available, so the missing value was not caused by an unavailable API property.

### Root Cause And Fix

- Backup: `_backups\system-routing-minimum-size-20260714-1115`.
- All three `SystemTypeDetailSummaryService` copies concatenated the criterion index and minimum token without a separator, producing legacy text such as `size 1min=0.328... max=0.984...`.
- The shared browser parser required a word boundary before `min`. A digit immediately before `min` prevented the match, while `max` still matched because it already had a preceding space. This exactly explains the asymmetric display.
- Changed all three writers to join the index, minimum, and maximum as separate tokens: `size 1 min=... max=...`.
- Relaxed only the shared minimum/maximum token parser so it accepts both the corrected form and already-persisted `size 1min=...` records. Existing snapshots therefore recover their minimum values without a mandatory rescan; future scans store the corrected representation.
- Strengthened the IE harness to verify minimum and maximum cells independently in `mm` and `in`, including unbounded sentinel values. The audit fixture intentionally keeps one malformed legacy criterion so backward compatibility remains covered.

### Verification

- Static/action contract checks: PASS for all three host trees (`77` generated actions, `251` exact routes, `59` prefix routes, `11` browser functions each).
- Focused Revit 2019 IE `WebBrowser` regression: `6/6 PASS`; legacy `size 1min=...` rendered minimum `100 mm`, maximum `300 mm`, and converted them independently to `3.937 in` / `11.811 in`.
- Full Release build and staged add-in verification: PASS for Revit 2019/2021/2023/2025/2027, zero compile errors.
- 2,000-row performance/cache gate: PASS.
- Full HTML/click/language/layout harness: `136 OK`, zero failures. Revit 2021 UI runtime is the expected `SKIP runtime-not-installed`; its build and stage verification passed.
- Final report: `artifacts\family-browser-ui-audit\20260714-routing-minimum-size-final\quality-gate-summary.md`.
- ProgramData installation and installer packaging were not requested and were not run (`Install: False`).

### Needs Revit Check

- Execute Next Work Queue item 20 with a real pipe or duct type. Compare one bounded Routing Preferences rule directly against the main and detached detail tables, confirming both minimum and maximum in `mm` and `in`.
- A rescan is not required to repair the display. A later precise scan is optional and only normalizes the persisted criterion text to the corrected spaced format.

## 2026-07-14 Routing Minimum Fix Five-Version Installer

### Build And Package

- Audit backup: `_backups\installer-routing-minimum-size-20260714-1229\FAMILY_BROWSER_BUTTON_AUDIT.md`.
- Rebuilt the current source and regenerated the Release Stage for Revit 2019/2021/2023/2025/2027 after the Routing Preferences minimum-size fix.
- Build result: zero compile errors. The existing 2025/2027 WindowsBase reference warnings remain warnings only.
- `Verify-FamilyBrowserRecovered.ps1` passed for all five staged add-in manifests and payloads.
- Existing installer artifacts were preserved; the package script's old-artifact cleanup was intentionally bypassed.
- Inno Setup 6.7.3 compiled all five version tasks and payloads successfully:
  - `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260714-1228_Setup.exe`
  - Size: `3,513,077 bytes` (`3.35 MB`)
  - SHA256: `717BE3F6981D4A65AF3165A8312215B1438248399CDEA9E4E3A665FF7BDDD318`
- The SHA256 was recomputed after compilation and matched the generated checksum. The compile log independently confirmed payload entries for `Rvt2019`, `Rvt2021`, `Rvt2023`, `Rvt2025`, and `Rvt2027`.
- Follow-up mail package backup: `_backups\mail-package-routing-minimum-20260714-1310\FAMILY_BROWSER_BUTTON_AUDIT.md`.
- Created the requested mail-sized package while preserving every previous package:
  - `artifacts\family-browser\mail-packages\20260714_02.zip`
  - Size: `15,990,939 bytes` (`15.25 MB`)
  - SHA256: `32BA50AF7DC6469ACEF44786E8E79419761E2CC37B2F052823F9BF2416D27DEF`
- ZIP validation passed for `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`. The SHA256 of `Setup.exe` read directly inside the ZIP matched the standalone installer: `717BE3F6981D4A65AF3165A8312215B1438248399CDEA9E4E3A665FF7BDDD318`.
- The temporary package work directory was removed after validation. ProgramData installation was not run.

## 2026-07-14 Missing Nested Family Relationship Classification

### Problem And Root Cause

- Fixture: the approved composite family `A` contains nested families `a`, `b`, and `c`, while the current-project copy of `A` contains only `a` and `b`.
- Existing relationship propagation correctly detected that `c` was absent and marked `A` as different, but the missing child row inherited the generic `DifferentFromStandard` state.
- `HasProjectFingerprintMissing` then treated any row with a standard fingerprint and an empty project fingerprint as a capture failure. A family that does not exist in the current project therefore appeared as `현재 프로젝트 Fingerprint 생성 실패`, even though no capture should have been attempted.

### Fix

- Backup: `_backups\nested-family-missing-relation-20260714-133650`.
- Added explicit relationship states `NestedMissingFromParent` and `NestedExtraInParent` in the shared nested-family propagation service.
- A missing child such as `c` now appears as `하위 패밀리 누락` with memo `상위 패밀리에서 누락: A` and action `상위 패밀리 업데이트 필요`.
- The child remains a non-selectable helper row and cannot be loaded independently. The affected parent `A` remains `DifferentFromStandard`, preserving recursive propagation through higher composite-family levels.
- Fingerprint failure is now shown only when project-side evidence exists, such as a project fingerprint/signature, capture failure reason, type count, or instance count. A completely absent child is no longer classified as a failed fingerprint capture.
- Added symmetric handling for a nested child that exists only in the current parent composition: `추가 하위 패밀리` / `상위 패밀리 구성 검토`.
- Added the new states to comparison counts, filters, modeler ranking, action/status display, detail memo generation, and Korean/English translation paths in all three host trees.
- Reverted an exploratory parent-document capture fallback because it was broader than this relationship-classification defect and was not required for the fix.

### Verification

- Shared propagation behavior test: PASS for `A={a,b,c}` versus project `A={a,b}`, including blank project fingerprint/signature on `c`, explicit child relationship state, parent propagation, and project-only nested child handling.
- Static/action/contract checks: PASS for all three host trees (`77` generated actions, `253` exact routes, `59` prefix routes, `11` browser functions each).
- Release build and staged add-in verification: PASS for Revit 2019/2021/2023/2025/2027 with zero compile errors. Existing 2025/2027 framework-reference warnings remain warnings only.
- 2,000-row performance/cache gate: PASS.
- HTML/IE `WebBrowser` click, language, and layout harness: `136 OK`, zero failures; Revit 2021 UI runtime is the expected `SKIP runtime-not-installed` while its build and stage verification passed.
- The first full harness exposed two English translation-order regressions in the new relationship memo; both were fixed and the complete gate was rerun successfully.
- Final report: `artifacts\family-browser-ui-audit\20260714-nested-missing-relation-final3\quality-gate-summary.md`.
- ProgramData installation and installer packaging were not requested and were not run (`Install: False`).

### Needs Revit Check

- With a real approved `A={a,b,c}` and current-project `A={a,b}`, rerun the current-model precise check.
- Confirm that `c` shows `하위 패밀리 누락`, names `A` as the parent, does not show `Fingerprint 생성 실패`, and cannot be selected directly.
- Confirm that `A` is marked different and its detail identifies missing child `c`. This runtime check validates Revit-derived relationship input; the downstream classification and rendering paths are covered by automation.

## 2026-07-14 Family/System Detailed Filter Reset

### Problem And Root Cause

- Family Load and System Type Load share the browser-shell Detailed Filter state: `advStatus`, `advGroup`, `advCategory`, and `advMismatchOnly`.
- `resetWorkflowFilters(...)` returned immediately whenever the destination was either browser tab. During the host's in-place pane replacement, the previous tab's JavaScript globals therefore remained alive and leaked into the other list.
- Result: a non-`All` Detailed Filter selected in Families was still selected in System Types, and the reverse transition behaved the same way.

### Fix

- Backup: `_backups\browser-cross-tab-detail-filter-reset-20260714-1430`.
- Added shared `resetDetailedFilterState()` to reset the Detailed Filter model and its visible controls together:
  - status and group return to `All`
  - category text is cleared
  - mismatch-only is unchecked
  - an open Detailed Filter mask is closed
- `setTab(...)` now passes both the previous and next tab to `resetWorkflowFilters(...)`. The reset runs only for an actual `families <-> systems` transition.
- A same-tab redraw does not reset the Detailed Filter. Search text, trade/category tree selection, checked rows, paging state, and the separate inline status filter remain outside this reset.
- Extended the IE `WebBrowser` harness to seed a non-default Detailed Filter and verify both Family-to-System and System-to-Family transitions, including JavaScript state and DOM controls.
- The first focused harness assumed both panes existed in the DOM simultaneously. The production dashboard keeps only the active pane, so the test was corrected to mirror the host's in-place pane replacement before the final verification.
- Added static guards for the cross-browser condition, the reset helper, the visible `All` state, and the new harness assertion.

### Verification

- Static/action/contract checks: PASS for all three host trees (`77` generated actions, `253` exact routes, `59` prefix routes, `11` browser functions each).
- Focused Revit 2019 HTML/click regression after correcting the test seam: `34/34 PASS`, zero failures.
- Release build, Stage generation, and staged add-in verification: PASS for Revit 2019/2021/2023/2025/2027.
- 2,000-row performance/cache gate: PASS. Each installed runtime target rendered 1,000 Family and 1,000 System rows with 150 DOM rows per page; filter response was `13-16ms`.
- Full HTML/click/language/layout harness: `136 OK`, zero failures. Revit 2021 UI runtime is the expected `SKIP runtime-not-installed`; its build and Stage verification passed.
- Final report: `artifacts\family-browser-ui-audit\20260714-cross-tab-detail-filter-reset-final\quality-gate-summary.md`.
- ProgramData installation and installer packaging were not requested and were not run (`Install: False`).

### Optional Revit Confirmation

- In Families, set Detailed Filter to a non-`All` status/group/category or mismatch-only, then switch to System Types and confirm the Detailed Filter opens at `All`; repeat in the reverse direction.
- This transition is covered by the real IE `WebBrowser` harness, so the runtime check is optional visual confirmation rather than an unresolved code-path requirement.

## 2026-07-14 Detailed Filter Reset Five-Version Installer

### Build And Package

- Audit backup: `_backups\detail-filter-reset-installer-20260714-150451\FAMILY_BROWSER_BUTTON_AUDIT.md`.
- Rebuilt the latest source and regenerated Release Stage payloads for Revit 2019/2021/2023/2025/2027 after the Family/System Detailed Filter reset.
- Build result: zero compile errors. Revit 2025/2027 retain the previously documented `WindowsBase` framework-reference warnings only.
- `Verify-FamilyBrowserRecovered.ps1` passed all five staged add-in manifests and payloads.
- Existing installers and mail packages were preserved. The packaging script's old-artifact cleanup was intentionally bypassed.
- Inno Setup 6.7.3 compiled all five version tasks and payloads successfully:
  - `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_detail-filter-reset-20260714-1504_Setup.exe`
  - Size: `3,517,867 bytes` (`3.35 MB`)
  - SHA256: `0DAF0AE6A034D548935D7DBD7311253C8DC8865B893823748CAB68030EE41A9F`
- Created a new mail-sized package without replacing the prior packages:
  - `artifacts\family-browser\mail-packages\20260714_03.zip`
  - Size: `16,275,824 bytes` (`15.52 MB`)
  - SHA256: `84B5AC196E44C1F818CCDC64986AD0D2196830204ABDEB0029E4CF3F4A31ABAC`
- ZIP validation passed for `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`. The embedded `Setup.exe` SHA256 exactly matches the standalone installer.
- `latest-build.json` and `latest-mail-package.json` now point to these validated artifacts.
- ProgramData installation was not requested and was not run.

## 2026-07-14 Cold-Start Project Comparison Restore

### Problem And Root Cause

- After completing a standard precise scan and the comparison for project `A`, restarting Revit after installing a new Family Browser build showed standard-only rows instead of the saved comparison result.
- The installer does not delete or overwrite the homepage-managed policy, snapshots, comparison reports, project records, thumbnails, or local V2 cache. Restarting Revit merely exposed an existing cold-start defect.
- The two-stage startup intentionally deferred expensive UI-thread file reads. Its first row-cache key used the literal values `deferred-project-scan-cache` and `deferred`, which could not match the key used when the real project comparison had been saved.
- After that row-cache miss, the deferred path explicitly refused to read the original project comparison record and rebuilt standard-only rows. It then persisted those incomplete rows under the deferred key.

### Fix

- Backup: `_backups\cold-start-project-comparison-restore-20260714-152031`.
- The UI thread now captures project title, local model path, and worksharing central path before starting the background preload. Central identity is therefore stable across different PCs and local-file locations.
- Added a Revit-free project identity overload to `ProjectSnapshotStore`. Document-based and background loads now share the same validation core, including standard source/snapshot revision, standard RVT file stamp, project file timestamp/length, alias collision protection, and referenced JSON existence checks.
- Startup preload now reads the selected trade's saved `ProjectScanCacheLoadResult` on the worker thread and passes it to the initial dashboard render. No `Document` API is called from that worker.
- Prepared standard-list and project-comparison revisions are included in the persistent row-cache key. The key schema was advanced to `v6-persistent-v2-project-restore`, so previously persisted standard-only deferred caches cannot mask a valid comparison.
- A failed or stale comparison is not silently accepted. Its exact validation reason remains in the dashboard runtime diagnostics, while a valid saved report is used immediately on the first render.
- Added static regression contracts for central-path capture, background comparison preload, prepared-result consumption, cache revision keys, shared validation core, and removal of the old startup skip path.

### Verification

- Static/action/contract checks: PASS for all three host trees (`77` generated actions, `253` exact routes, `59` prefix routes, `11` browser functions each).
- Release build and staged add-in verification: PASS for Revit 2019/2021/2023/2025/2027 with zero compile errors. Existing 2025/2027 framework-reference warnings remain warnings only.
- 2,000-row performance/cache gate: PASS.
- HTML/IE `WebBrowser` click, language, detail, pending-save/sync, and layout harness: `136 OK`, zero failures. Revit 2021 UI runtime is the expected `SKIP runtime-not-installed`; its build and stage verification passed.
- Final report: `artifacts\family-browser-ui-audit\20260714-project-comparison-cold-start-restore-final\quality-gate-summary.md`.
- Homepage bootstrap audit reached `https://update.zerokky.com/family-browser/bootstrap.json` successfully, but this PC could not access either configured managed-path candidate (`I:\30. 협력사 전용폴더\00. BIM_KCIM\02. 패밀리\TEST`, `D:\TEST`). The audit therefore recorded `UNAVAILABLE` with zero structural failures; project `A`'s actual managed files must be confirmed on the connected test PC.

### Needs Revit Check

- On the PC that can access the managed folder, open the already-compared project `A` after installing the corrected build. Do not rerun the precise comparison first.
- Confirm that Family and System Type rows restore the prior comparison states on the initial browser load. A manual Refresh must not be required.
- If the saved comparison is legitimately stale because the project file or approved standard changed after the scan, confirm that Debug Log reports the exact stale timestamp/length/source reason instead of silently showing a valid comparison.

## 2026-07-14 Cold-Start Restore Five-Version Installer

### Build And Package

- Reused the quality-gate-verified Release Stage for Revit 2019/2021/2023/2025/2027; `Verify-FamilyBrowserRecovered.ps1` had passed all five manifests and payloads in the final gate.
- Existing installer and mail-package artifacts were preserved.
- Inno Setup 6.7.3 compiled all five version tasks and payloads successfully:
  - `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_project-comparison-restore-20260714-1548_Setup.exe`
  - Size: `3,517,710 bytes` (`3.35 MB`)
  - SHA256: `4C40678B37A10BB1CB409BF87E21C9D8F9A90F2F6D3F95964C65645829D84138`
- Created a new mail-sized package without replacing prior packages:
  - `artifacts\family-browser\mail-packages\20260714_04.zip`
  - Size: `16,357,981 bytes` (`15.60 MB`)
  - SHA256: `4E68501B8D595E3DEEC0EB1B81E7A40F5124CB66C7EE5D9416EA8617D405137F`
- ZIP validation passed for `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`. The embedded installer SHA256 exactly matches the standalone installer.
- `latest-build.json` and `latest-mail-package.json` point to these validated artifacts. ProgramData installation was not run.

## 2026-07-14 Duplicate Search Status Filter Removal

### Problem And Scope

- Family Load and System Type Load rendered two status-filter groups: an always-visible count row directly below search and the context-aware `State` filters beside Load/Apply Selected Items.
- The search-row group duplicated the same workflow and continued to show zero-count buttons, consuming vertical space without adding a distinct action.
- The requested change removes only the search-row summary filters. The Detailed Filter dialog link, trade filters, and inline `State` filters remain available.

### Fix

- Backup: `_backups\remove-search-summary-filters-20260714-155520`.
- Removed the `AppendFilterBar(sb)` call from `AppendBrowserSearchChrome(...)` in all three Revit host trees.
- Preserved `AppendDisciplineFilterBar(...)`, `AppendFamilyInlineStatusFilterBar(...)`, and `AppendSystemInlineStatusFilterBar(...)` unchanged.
- Reduced the browser search chrome's fallback height and made the final height follow its actual `scrollHeight` within `84-140px`. Long trade labels can therefore wrap without clipping while the removed row no longer leaves a blank band.
- Added static and IE `WebBrowser` harness guards that fail if a `.filterbar` summary row is rendered again.

### Verification

- Static/action/contract checks: PASS for all three host trees (`77` generated actions, `253` exact routes, `59` prefix routes, `11` browser functions each).
- Release build, Stage generation, and staged add-in verification: PASS for Revit 2019/2021/2023/2025/2027 with zero compile errors. Existing framework/platform analyzer warnings remain warnings only.
- 2,000-row performance/cache gate: PASS. Search/filter response was `10-15ms`; each 1,000-row list kept 150 DOM rows per page.
- HTML/IE `WebBrowser` click, language, detail, and layout harness: `136 OK`, zero failures. Revit 2021 runtime remains the expected `SKIP runtime-not-installed`; its build and Stage verification passed.
- Generated dashboard HTML checked: `120` files, `.filterbar` summary markup found in `0` files.
- Visual 1280px Family render checked: search, Detailed Filter, visible count, and trade row remain; the inline `State` filters remain beside the selection actions with no grid overlap.
- Final report: `artifacts\family-browser-ui-audit\20260714-remove-search-summary-filters-final\quality-gate-summary.md`.
- ProgramData installation and installer packaging were not requested and were not run (`Install: False`).

## 2026-07-14 Search Filter Cleanup Five-Version Installer

### Build And Package

- Backup: `_backups\installer-search-filter-cleanup-20260714-170007`.
- Reused the Release Stage produced by the successful `20260714-remove-search-summary-filters-final` quality gate and reran staged add-in verification.
- `Verify-FamilyBrowserRecovered.ps1` passed Revit 2019/2021/2023/2025/2027 manifests and payloads.
- Existing installers were preserved; the old-artifact cleanup routine was not invoked.
- Inno Setup compiled all five version tasks and payloads successfully:
  - `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_search-filter-cleanup-20260714-1700_Setup.exe`
  - Size: `3,518,662 bytes` (`3.36 MB`)
  - SHA256: `46FF853F5D0FA236EE5786154BFC914C390AC5F69AE4BE7C9FFA29B7D7CEF47E`
- Installer validation passed: PE `MZ` header, recomputed SHA256, checksum file, and `latest-build.json` all match.
- The Inno compile log independently contains Stage payloads for `Rvt2019`, `Rvt2021`, `Rvt2023`, `Rvt2025`, and `Rvt2027` and ends with `Successful compile`.
- A mail-sized ZIP and ProgramData installation were not requested and were not created.

### Mail Package

- Backup: `_backups\mail-package-search-filter-cleanup-20260714-170513`.
- Created a new package without replacing any prior mail package:
  - `artifacts\family-browser\mail-packages\20260714_05.zip`
  - Size: `16,515,222 bytes` (`15.75 MB`)
  - SHA256: `43C79E6FF8A5EFFD614AB0781458B77094224A78A8563D405BFB48135F8E5EE3`
- ZIP entries validated: `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`.
- The `Setup.exe` SHA256 read directly inside the ZIP matches the standalone installer: `46FF853F5D0FA236EE5786154BFC914C390AC5F69AE4BE7C9FFA29B7D7CEF47E`.
- `latest-mail-package.json` and the ZIP checksum file match the generated package. Temporary packaging files were removed after validation.

## 2026-07-15 File Guard Runtime Enforcement And Diagnostics

### Confirmed Causes

- This PC cannot reach either homepage-managed candidate: `I:\30. 협력사 전용폴더\00. BIM_KCIM\02. 패밀리\TEST` or `D:\TEST`. File Guard is shared policy, so the screenshot failure occurred before a policy could be saved.
- The error dialog then obscured the cause because diagnostics attempted the same unavailable managed root and returned `(managed log unavailable)`.
- Admin OFF wrote personal UI state to the managed share, rebuilt the full browser shell, and traversed the full Autodesk ribbon object graph. Those synchronous operations explain the long delay seen on the other PC.
- Runtime guard contexts used the open local RVT path/title but omitted the worksharing central path on fast checks. A differently named local copy could therefore fail to match the guarded central RVT and allow Load Family or type edits.

### Fix

- Backup: `_backups\file-guard-runtime-enforcement-20260715-111457`.
- Personal language/Admin state now lives under `%LOCALAPPDATA%\KKY\FamilyBrowser\Settings`; shared File Guard policy remains in the managed folder.
- File Guard configuration now checks managed-root availability before opening/saving and shows a specific shared-folder message. Error logs fall back to `%LOCALAPPDATA%\KKY\FamilyBrowser\Diagnostics\Errors` when the managed log root is unavailable.
- Homepage bootstrap now stops before policy writes when all managed candidates are unavailable.
- Admin toggle hands the already-loaded policy directly to the native guard, renders the active pane without a full data reload, and no longer recursively traverses Autodesk ribbon controls.
- Native guard contexts cache and include the worksharing central model path, so a local copy can match a File Guard entry registered against its central RVT.
- The UI audit harness now asserts Admin OFF blocks `EditFamilies` and `AddDeleteTypes` for a differently named local copy matched through its central path, while Admin ON and non-target projects remain allowed.
- Removed `Application.ExitThread()` from the audit harness shutdown path after the gate exposed completed scenarios waiting unnecessarily before process exit.

### Verification

- Static/action/contract checks: PASS for all three host trees (`77` generated actions, `253` exact routes, `59` prefix routes, `11` browser functions each).
- Release build and Stage verification: PASS for Revit 2019/2021/2023/2025/2027.
- 2,000-row performance/cache gate: PASS.
- HTML/IE `WebBrowser`, language, layout, message, and File Guard policy harness: `136 OK`, zero failures for installed runtimes. Revit 2021 remains the expected `SKIP runtime-not-installed`; its build and Stage verification passed.
- Final report: `artifacts\family-browser-ui-audit\20260715-file-guard-runtime-final-v2\quality-gate-summary.md`.
- ProgramData installation was attempted after confirming `Revit.exe` was closed, but Windows denied overwrite of the existing administrator-owned 2019 DLL. The installer stopped on its first copy, so the installed 2026-07-13 payload was left unchanged. Use an elevated installer/install shell for deployment. Installer packaging was not requested and was not run.

### Needs Revit Check

- Follow Next Work Queue item 23 on the PC that can reach the homepage-managed folder. The remaining check is actual Revit native command cancellation, not policy matching or browser routing.

## 2026-07-15 File Guard Path Alias Normalization

### Finding

- `Path.GetFullPath()` normalized separators and relative segments but did not make `I:\...`, `\\BIM-SERVER\share\...`, and `\\10.20.30.40\share\...` the same path.
- A file-name fallback masked some cases, but a differently named workshared local copy still depended on resolving and matching the central path. This made path spelling a real remaining weakness.
- The Admin ON/OFF delay was not expected behavior. It was a regression from managed-share user-setting writes, full dashboard/data regeneration, and recursive Autodesk ribbon traversal; those operations were removed in the preceding fix.

### Fix And Verification

- Backup: `_backups\file-guard-path-alias-20260715-115806`.
- Added cached Windows `WNetGetConnection` expansion for mapped-drive paths. No drive enumeration or DNS lookup is performed on every guarded command.
- Added endpoint-neutral UNC identity using the exact `share + subfolder + RVT filename`, allowing hostname/IP server aliases while keeping different shares isolated.
- Extended the runtime-policy harness so the guarded target `\\BIM-SERVER\bim\CentralModel.rvt` blocks a differently named local copy whose Revit central identity is `\\10.20.30.40\bim\CentralModel.rvt`.
- Static/action/contract checks: PASS for all three host trees.
- Revit 2019/2021/2023/2025/2027 Release build and Stage verification: PASS with zero compile errors.
- Focused 2019 compiled-host alias test: PASS, zero failures/warnings; hostname/IP same-share match and different-share isolation both passed.
- ProgramData still contains the older administrator-owned DLL because the current non-elevated shell cannot overwrite it. Runtime testing requires the next elevated installer build.

## 2026-07-15 File Guard Family/Type Command Enforcement

### Recheck Result

- Revit native Load Family is bound to `EditFamilies`; the database-level `FamilyLoadingIntoDocument` event independently cancels API/native family loads when the target file has `BlockFamilyLoadAndEdit` enabled.
- Family/type Rename is bound to `RenameFamilyOrType`, which deliberately requires both `EditFamilies` and `AddDeleteTypes`. The protected-content updater separately distinguishes `Family` from `FamilySymbol` / `ElementType`, posts an error failure, and forces transaction rollback before commit.
- Approved Family Browser family load and system-type apply paths remain inside `BeginTrustedOperation(...)`, so the guard blocks direct Revit operations without blocking the controlled browser workflow.
- A remaining activation gap was found: command and database-event guards started with Revit, but the transaction-level updater could wait for an Admin toggle or policy-save refresh before registration. This left non-command/API type mutations without the final rollback layer immediately after opening a protected document.

### Fix And Verification

- Backup: `_backups\file-guard-command-enforcement-20260715-121451`.
- `NotifyActiveDocumentChanged(...)` now refreshes protected-change updater registration after caching the central identity. The updater is therefore ready on the first active protected document and still avoids a full family/type baseline scan.
- Extended static guards for Load Family command binding, family-load database event attachment, combined Rename mapping, updater rollback, active-document registration, and trusted Family Browser load/apply scopes.
- Extended the compiled-host harness to verify independent family/type flags, Admin OFF/ON rename decisions, `LoadFamily` command presence, and required-permission wiring.
- Full quality gate PASS: static/contract, nested-family propagation, Revit 2019/2021/2023/2025/2027 Release build and Stage verification, 2,000-row performance/cache gate, and `136` IE `WebBrowser` scenarios with zero failures. The `24` harness warnings were expected handled alerts for clicking Load/Apply without a selected row.
- Final report: `artifacts\family-browser-ui-audit\20260715-file-guard-command-enforcement\quality-gate-summary.md`.
- ProgramData was not changed. The installed DLL is still the older administrator-owned payload; use the next elevated installer before runtime validation.

### Needs Revit Check

- Next Work Queue item 23 remains required: in a reachable managed-folder environment, verify native Load Family cancellation and family/type rename rollback on both the central RVT and a differently named local copy, then confirm Admin ON restores both operations.

## 2026-07-15 Shared KKY Tools Ribbon Tab And Family Browser Icon

### Finding And Fix

- Backup: `_backups\family-browser-shared-ribbon-tab-20260715-124549`.
- All three Family Browser host applications created and queried a separate `KKY Browser` ribbon tab, while KKY Tool already used `KKY Tools`. This forced the two products into different Revit tabs.
- Replaced the Family Browser tab constant with `KKY Tools` and changed tab creation, panel lookup, and panel creation to use `TabName` / `PanelName` consistently. Both add-ins can now load in either order and reuse the same Revit tab; KKY Tool remains in its `Hub` panel and Family Browser remains in its `Family Browser` panel.
- Added a dedicated navy/blue Family Browser library-folder icon at native Revit sizes: `family-browser-ribbon-16.png` and `family-browser-ribbon-32.png`.
- Added `New-FamilyBrowserRibbonIcons.ps1` as the deterministic icon source and `FamilyBrowserRibbonIcon.cs` as the shared embedded-resource loader. The button now receives its tooltip, small image, and large image in all three hosts.
- Added WPF `ImageSource` support and embedded both PNG resources into the 2019-2023, 2025, and 2027 host projects.
- Static QA now rejects the retired `KKY Browser` tab name and requires the shared tab, separate Family Browser panel, icon bindings, exact 16/32 PNG dimensions, WPF support, and embedded logical resource names.

### Verification

- Static/action/contract and nested-family propagation checks: PASS for all three host trees (`77` generated actions, `253` exact routes, `59` prefix routes, `11` browser functions each).
- Release build: PASS with zero compile errors for the 2019/2021/2023 shared host, 2025 host, and 2027 host.
- Revit 2019/2021/2023/2025/2027 Stage generation and manifest/payload verification: PASS.
- Compiled DLL resource smoke: PASS. The 2019-family, 2025, and 2027 assemblies each exposed both embedded resources; the runtime loader returned frozen `16x16` and `32x32` images. The 2019, 2021, and 2023 staged DLL hashes are identical as intended.
- 2,000-row performance/cache gate: PASS. Shell `3-11ms`, cold usable `448-685ms`, warm usable `296-389ms`, filter `11-15ms`.
- Quality summary: `artifacts\family-browser-ui-audit\20260715-shared-ribbon-tab-icon\quality-gate-summary.md` (`SkipHarness`; all recorded steps OK).
- A full long-running IE harness was also started. It completed `112` scenarios with zero failures and `18` handled-alert warnings before the outer 10-minute command limit interrupted the remaining `24` Revit 2027 scenarios. No dashboard HTML/JS changed in this patch; the last complete baseline remains `136` scenarios with zero failures from `20260715-file-guard-command-enforcement`.
- ProgramData and installer artifacts were not changed in this task.

### Needs Revit Check

- After elevated deployment, start Revit with both KKY Tool and Family Browser enabled. Confirm there is one `KKY Tools` tab, separate `Hub` and `Family Browser` panels, and a sharp Family Browser icon at normal/high-DPI ribbon sizes.

## 2026-07-15 Global Overflow Tooltip And Managed-Folder First-Use Recovery

### Problems Found

- The previous dashboard `applyControlTitles()` covered only selected control classes and could assign titles without checking whether text was actually clipped. It also did not reliably remove a generated title after a user widened a resizable table column, and auxiliary HTML surfaces did not share the same behavior.
- When all managed-folder candidates received from the homepage were missing or unreachable, startup could continue with only readiness diagnostics. A first-time user did not get a clear Home recovery flow, and a management-only workstation had no guarded way to select an internal shared folder as its managed root.
- The queued-search focus audit could inherit an IE `propertychange` timer and virtual-row filter signature from a preceding check. This caused an intermittent false failure in otherwise valid pending-save/sync scenarios.

### Fix

- Backup: `_backups\family-browser-overflow-managed-folder-20260715-131734`.
- Added shared `FamilyBrowserOverflowTitleScript` with delegated IE-compatible hover/focus handling. It compares `scrollWidth/clientWidth` and `scrollHeight/clientHeight`, uses `data-full-text` where supplied, preserves authored `title` text, and removes only its own generated title as soon as the element is no longer clipped.
- Applied the shared overflow service to the main dashboard, detached detail, request composer, Family selection, Standard RVT manager, File Guard, standard-list sheet selection, and every common full-HTML message/result/confirmation window. Resizable Family/System columns explicitly refresh clipping state after drag completion.
- Added `FamilyBrowserManagedFolderSetupService`. A valid persisted manual root is restored before homepage bootstrap; otherwise the first ready browser shell checks whether the homepage policy resolves to a usable managed folder.
- When no registered network managed-folder path is reachable, a structured Korean/English HTML notice shows the requested manager/distributor message, separates general-user guidance from management-only setup, and warns that generated files/folders must not be edited, moved, renamed, or deleted manually.
- Management-only setup accepts only an existing writable UNC path or mapped Windows network drive. It rejects local folders, performs a write probe, creates the required managed hierarchy and `KKY_FAMILY_BROWSER_MANAGED_FOLDER_README.txt`, atomically stores only a per-user pointer under `%LOCALAPPDATA%`, and refreshes the active policy/UI immediately.
- Stabilized the queued-search audit by cancelling stale IE timers, resetting row-window page/render/filter signatures, and searching with a guaranteed existing item name. The formerly intermittent 2025 pending-system scenario passed `5/5` repeated runs.

### Verification

- Static/action/contract and nested-family propagation checks: PASS for all three host trees (`77` generated actions, `253` exact routes, `59` prefix routes, `11` browser functions each).
- Release build: PASS with zero compile errors for the 2019/2021/2023 shared host, 2025 host, and 2027 host. Revit 2019/2021/2023/2025/2027 Stage generation and manifest/payload verification passed.
- 2,000-row performance/cache gate: PASS; measured search/filter response was `11-17ms` and all performance result failures were zero.
- Full IE `WebBrowser` HTML/click/language/layout/detail/overflow/managed-folder harness: `136/136 PASS`, zero failures. The `25` warnings were expected handled alerts from intentionally clicking Load/Apply/Detail with no selected row. Revit 2021 runtime remains `SKIP runtime-not-installed`; its build and Stage package passed.
- Quality report: `artifacts\family-browser-ui-audit\20260715-overflow-managed-folder-final\quality-gate-summary.md`.
- Full harness report: `artifacts\family-browser-ui-audit\20260715-overflow-managed-folder-harness-final\ui-harness-summary.md`.
- ProgramData installation and installer packaging were not requested and were not run.

### Needs Revit Check

- Follow Next Work Queue item 24 on a first-use workstation. Verify native IE tooltip timing, real column drag behavior at workstation DPI, and the exact first-open message flow.
- Use a disposable internal UNC or mapped network share for the setup test. Confirm the pointer survives a Revit restart and the generated hierarchy/readme are correct. Automated tests deliberately did not write to a real company managed share.

## 2026-07-15 Family Edit Dialog Guard And Scan Warning Export

### Problem And Cause

- Standard precise scan, nested-family scan, thumbnail capture, and Current Model Check already placed `EditFamily` work inside `FamilyThumbnailConstraintDialogGuard`, but automatic handling depended mainly on a known warning-text list.
- `Opening not cutting anything` was not in that list. More generally, a new Revit family-edit warning could expose a safe enabled OK button and still remain open because the guard did not use the actual native button topology as its primary safety signal.
- Delete Instance/Delete Type/Delete dialogs need the opposite rule: when no safe OK/Confirm/Continue exists and the only enabled choices are destructive actions plus Cancel, the scan must choose Cancel and never click a destructive action.
- Auto-handled dialog records were counted, but the standard-scan summary did not identify affected families and Current Model Check Excel export did not include the full family-edit warning history.

### Fix

- Backup: `_backups\family-edit-dialog-guard-20260715-01`.
- Updated all three host implementations of `FamilyThumbnailConstraintDialogGuard`.
  - While a family-edit scope is active, any actual enabled `OK`, `Confirm`, or `Continue` native button now wins regardless of previously unseen warning text.
  - Added an explicit `OpeningNotCuttingAnything` reason for the reported warning.
  - If no safe confirmation exists and enabled buttons consist only of Delete/Remove actions plus Cancel, the guard clicks Cancel.
  - Unknown Cancel-only dialogs and dialogs outside an active family-edit scope remain untouched. This avoids approving unrelated Revit prompts.
  - Removed the old confirm fallback that could accidentally target Cancel, broadened native button-class recognition, synchronized current family context, and reduced polling from 250ms to 150ms.
- Standard-scan result summaries now list up to five affected `category / family / action / reason` records instead of showing only a count.
- Current Model Check exposes Excel export when either comparison rows or handled family-edit dialogs exist. Export remains user-triggered through Save As; no workbook is created automatically.
- The workbook adds a localized `스캔경고` / `ScanDialogs` sheet containing UTC time, category, family, action, reason, override result, enabled buttons, and full dialog text.
- Added a public audit seam and compiled-host tests for unseen OK warnings, the reported opening warning, OK precedence over Delete, delete-only Cancel, inactive scope, and unrelated Cancel-only dialogs. The harness also creates and inspects a real temporary XLSX, then deletes it.

### Coverage

- Covered paths include Standard RVT full/precise/selected/nested family scans, family thumbnail/edit-document capture, and Current Model Check deep family capture because each path sets the current family context around `EditFamily` and passes the same guard instance.
- Family/System load or apply operations were not broadened by this change; they are separate mutation workflows and should not automatically dismiss arbitrary Revit dialogs.

### Verification

- Static/action/contract and nested-family propagation checks: PASS for all three host trees (`77` generated actions, `253` exact routes, `59` prefix routes, `11` browser functions each).
- Release build and Stage generation/verification: PASS with zero compile errors for Revit 2019/2021/2023/2025/2027.
- Focused compiled-host guard/XLSX audit on the 2019-family, 2025, and 2027 assemblies: PASS, zero failures. This includes a destructive `Delete Instance` button deliberately carrying `IDOK(1)`; it still resolves to Cancel.
- 2,000-row performance/cache gate: PASS.
- Full IE `WebBrowser` HTML/click/language/layout/detail harness before the final control-ID safety tightening: `136/136 PASS`, zero failures. The safety tightening changed only dialog classification/audit code; post-change static/contract, five-target build/Stage, performance, and three compiled-host focused audits all passed. Revit 2021 remains the expected `SKIP runtime-not-installed`; its build and Stage package passed.
- Full harness report: `artifacts\family-browser-ui-audit\20260715-family-edit-dialog-final\quality-gate-summary.md`.
- Post-safety build/Stage/performance report: `artifacts\family-browser-ui-audit\20260715-family-edit-dialog-safety-final\quality-gate-summary.md`.
- Post-safety compiled-host audit results: `artifacts\family-browser-ui-audit\20260715-family-edit-dialog-safety-focused`.
- ProgramData installation and installer packaging were not requested and were not run.

### Needs Revit Check

- Execute Next Work Queue item 25. Native Revit button discovery cannot be fully proven outside the Autodesk dialog runtime, especially if a future warning uses a nonstandard WPF window without discoverable child buttons.

## 2026-07-15 Family And System List Trade Switching

### Problem And Root Cause

- With Architecture and Mechanical both prepared, selecting Mechanical in Standards correctly produced Mechanical rows. Clicking the ready Architecture chip below the Family/System search field changed the target key but could spend a long time loading and leave the visible rows looking unchanged.
- Prepared startup data already contained registered target slots, but `FindPreparedSlotData(...)` accepted them only during the one initial preload render. An explicit list-target switch therefore ignored the ready Architecture snapshot/list and returned to source/network reads.
- Only the startup-selected trade's project comparison was preloaded. An alternate trade could receive a synthetic project-scan miss even when a saved comparison existed under the managed Projects store.
- The persistent row-cache key was calculated before that alternate saved comparison could be restored, so a cached no-comparison row set could win. The prior target's DOM also remained visible throughout the host refresh.

### Fix

- Backup: `_backups\family-browser-trade-switch-20260715-153052`.
- Added a separate, narrowly scoped `allowPreparedSlotData` refresh permission. Explicit trade switching may reuse prepared standard registration/snapshot/list data, while policy and other mutable state still come from the live store.
- `SetBrowseDiscipline(...)` now uses the progress-aware refresh path, identifies the selected target in the progress text, and enables prepared-slot reuse only for that refresh.
- Added on-demand `ProjectSnapshotStore.TryLoadLatestProjectScan(...)` restoration for the selected trade, including deployment identity, source registration, standard snapshot, and prepared-slot cache stamp.
- Added `PrimePreparedProjectScanForSelectedSlot(...)` before `BuildModelerAllSlotsCacheKey(...)`, preventing a stale no-comparison row cache from being chosen before the saved trade comparison is known. The row-cache generation is now `v7-trade-switch-project-restore`.
- Added IE-compatible `beginBrowseDisciplineSwitch(...)`. The clicked chip becomes active and old-target rows are hidden immediately while C# loads and replaces the pane; normal search/filter/detail behavior remains unchanged after the new rows arrive.
- Standard RVT registration/rescan/reset and approved-list connect/clear now discard `_startupPreloadResult`, so a later target switch cannot reuse prepared source data invalidated by those mutations.

### Automated Coverage

- Static QA locks the separate prepared-slot flag and `finally` reset, live-policy behavior, progress refresh, on-demand project restore, cache-stamp update, project-scan priming order, V7 cache key, immediate JS handler, mutation invalidation, and generated chip `onclick`.
- The IE `WebBrowser` harness now opens Family/System data scenarios with at least two trade chips, clicks an inactive chip, and verifies the target key/body marker, exactly one active chip, and zero visible rows belonging to the prior trade before restoring the original state.
- The new switch test runs after existing detail/filter/click checks so its state change cannot invalidate unrelated scenario expectations.

### Verification

- Static/action/contract and nested-family propagation: PASS for all three source host trees (`77` generated actions, `253` exact routes, `59` prefix routes, `11` browser functions each).
- Revit 2019/2021/2023/2025/2027 Release build and Stage manifest/payload verification: PASS with zero compile errors.
- 2,000-row performance/cache gate: PASS.
- Full HTML/IE `WebBrowser` click, language, layout, detail, and target-switch harness: `136 OK`, zero failures. The `25` warnings are expected handled alerts from intentionally clicking unavailable/unselected actions.
- Revit 2021 UI runtime: expected `SKIP runtime-not-installed`; its build and Stage package passed.
- Final report: `artifacts\family-browser-ui-audit\20260715-trade-switch-final\quality-gate-summary.md`.
- ProgramData installation and installer packaging were not requested and were not run.

### Needs Revit Check

- Execute Next Work Queue item 26 with two real ready trades. The automated harness proves target/DOM/cache routing; it cannot prove the company's live managed-folder contents or visually compare the resulting real Architecture and Mechanical family/type catalogs.

## 2026-07-15 Revit-Style Compound Layer Table And Persisted Units

### Problem And Cause

- Compound-layer data was captured with function, material, and thickness, but `SystemTypeDetailSummaryService` flattened every layer into one `@row` string. The detail UI therefore showed loose text rather than the ordered assembly users recognize from Revit Edit Assembly.
- Core boundaries, structural material, and variable-layer state were not part of `StandardSystemTypeLayerSnapshotItem`, so the renderer could not identify those semantics.
- The existing Routing Preferences unit dropdown always started at `mm`, affected only routing criteria, and forgot the selection when the browser closed.

### Fix

- Backup: `_backups\system-layer-unit-ui-20260715-160848` (26 existing source/audit files).
- Final visual-QA backup: `_backups\system-detail-hidden-block-qa-20260715-165016`.
- `CaptureSystemTypeLayers(...)` now records first/last core layer, structural material index, and variable layer index through a reflection-safe cross-version resolver. The snapshot model and comparison clone preserve `IsCore`, `IsStructuralMaterial`, and `IsVariable`.
- `SystemTypeDetailSummaryService` now emits a structured `@layer` row containing index, function, material, raw internal feet, display text, and the three semantic flags. It also keeps the legacy `@row` during migration.
- Shared IE-compatible rendering now shows an exterior-to-interior composition table with Revit-like columns, core boundary bands, structural/variable badges, and an internal horizontal scroll region for narrow detail windows.
- Legacy snapshots without `@layer` are parsed back from their existing `Function / Material / Thickness` rows. They remain readable, but core/structural/variable metadata appears only after a new precise scan.
- Routing criteria and layer thickness now use one synchronized display unit. First use defaults to `mm`; `in` converts from raw internal feet, not from rounded display text.
- The preference is atomically stored at `%LOCALAPPDATA%\KKY\FamilyBrowser\Settings\measurement-unit.txt`. Main and detached detail windows exchange `kkyfb:measurement-unit/*` and `setSystemDisplayUnitFromHost(...)` updates without rebuilding HTML or losing the current detail state.
- Final visual QA found that the audit-only detached-detail canvas forced every `.detail-block` visible. That test styling exposed the otherwise hidden Family composition and legacy bottom preview blocks in a System detail capture. The override was removed, production detached HTML now marks `@system-detail-v1` content with `fb-system-detail`, and both irrelevant blocks are hidden with an explicit `!important` guard.

### Automated Coverage

- Contract/static QA requires the new action prefix, LocalAppData preference service, host and detached routing, structured capture fields, raw `G17` feet, legacy parser, layer table, direction/core/badge markup, and synchronized selectors in all three host trees.
- The compiled-host audit verifies missing preference -> `mm`, saved `in` -> restored `in`, and unsupported `cm` -> normalized `mm` using a temporary file, so the user's real preference is not changed by QA.
- The IE `WebBrowser` fixture renders three layers (`100 mm`, `200 mm`, `15 mm`), requires two core boundaries, material rows, direction labels, and structural/variable badges, then verifies synchronized conversion to `3.937 in`, `7.874 in`, and `0.591 in` and restoration to `mm`.
- The harness now checks computed visibility for the legacy preview and Family composition blocks and fails if either appears in System detail. The wrapper also exports System detail HTML/PNG for each installed runtime target, not only Family detail.

### Verification

- Focused Revit 2025 IE harness: `34/34 OK`, zero failures in Korean/English and all dashboard scenarios.
- Full quality gate: static/contract and nested-family propagation PASS; Revit 2019/2021/2023/2025/2027 Release build, Stage generation, and Stage verification PASS; 2,000-row performance/cache PASS; `136 OK`, zero failures, plus expected Revit 2021 `SKIP runtime-not-installed`.
- Performance remained within targets: shell `3-13ms`, cold usable `452-783ms`, warm usable `304-444ms`, search/filter `11-17ms` for each 1,000-row tab.
- Final report: `artifacts\family-browser-ui-audit\20260715-system-layer-unit-final-v2\quality-gate-summary.md`.
- Full-height 2025 visual: `artifacts\family-browser-ui-audit\20260715-system-layer-unit-final-v2\harness\Rvt2025-admin-system-with-data-ko-light\detail-preview-full.png`.
- Stage DLL SHA256: 2019/2021/2023 `86E469E8186B86D91D82BD38CF8F3656E09B9EA39B5A5C32D2874667D1BC4288`; 2025 `E70542DBC2D00AF883AD0C823ED672CB6F3D5778639E2AB25469ABB8453AF0EF`; 2027 `F98E659915870993F37A69894567CF7080C6421B65816DD916A923B203C6146D`.
- ProgramData installation and installer packaging were not requested and were not run (`Install: False`).

### Needs Revit Check

- Execute Next Work Queue item 27. The automated fixture proves schema, rendering, persistence, conversion, and compatibility, but real Autodesk `CompoundStructure` content and exact visual order must still be compared with a scanned wall/floor/roof in Revit.

## 2026-07-15 Nested-Only Family Standalone Placement Guard

### Intent And Definition

- User policy: a family used only as a child of other families must not be modeled as a standalone project instance when the per-file restriction is enabled.
- `Nested-only` means the selected standard's latest precise scan records at least one parent reference and zero project-level instances whose `FamilyInstance.SuperComponent` is null.
- Category plus family name identifies the restriction target. A same-name family in another category is not blocked.
- Parent placement is allowed. Nested instances created underneath the parent are skipped because they have a `SuperComponent`.
- Admin Mode ON and a disabled per-file option allow direct placement. Existing standalone instances are not deleted retroactively.
- Legacy or incomplete snapshots fail open. A new precise scan is required before the policy can enforce safely.

### Implementation

- Backup: `_backups\nested-only-placement-guard-20260715-172727`.
- Dedicated policy: `FAMILY_BROWSER_NESTED_ONLY_PLACEMENT_POLICY.md`.
- Added per-target `BlockNestedOnlyStandalonePlacement` persistence, HTML checkbox, live refresh signature, summary count, and XLSX policy export field.
- Standard snapshot schema `6` captures `StandalonePlacementUsageCaptured` and `StandaloneInstanceCount` while excluding instances with a parent component.
- `FamilyBrowserNestedOnlyPlacementCatalogStore` publishes `<snapshot>.nested-only-placement-v1.json` and keeps a validated local cache under `%LOCALAPPDATA%\KKY\FamilyBrowser\Cache\Guard\NestedOnly`.
- Full and selected standard scan attachment calls `NotifyStandardSnapshotChanged()` after the new sidecar exists, so a previous negative lookup cannot delay enforcement for the former 30-second cache window.
- The existing Revit updater now listens for added `FamilyInstance` elements. With Admin OFF and the option enabled, a catalog match without `SuperComponent` posts the File Guard failure before commit and records `NestedOnlyFamilyPlacement`, `FamilyIsNestedOnlyInSelectedStandard`, and the known parent family names.
- Existing family load/edit and type-change restrictions remain independent. The generic `표시 항목 적용` bulk action deliberately does not enable this stricter option automatically.

### Automated Coverage

- Static QA checks the requested Korean/English label, six-field policy payload, persistence/clone/export paths, sidecar publishing, addition trigger, parent exemption, Admin bypass, rollback audit reason, and all three host trees.
- Semantic harness cases cover nested-only inclusion, dual-use exclusion, category mismatch exclusion, legacy metadata fail-open, Admin OFF block, and Admin ON bypass.
- Full report: `artifacts\family-browser-ui-audit\20260715-nested-only-placement-final-v3\quality-gate-summary.md`.
- Static/contract and nested-family propagation: PASS.
- Revit 2019/2021/2023/2025/2027 Release build and Stage verification: PASS.
- 2,000-row performance/cache gate: PASS. Shell `3-12ms`, cold usable `445-675ms`, warm usable `282-386ms`, filter `11-21ms` for 1,000-row tabs on installed runtimes.
- IE WebBrowser UI/click/language/layout harness: `136 OK`, zero failures, plus expected Revit 2021 `SKIP runtime-not-installed`.
- A v2 run received external keyboard text while the hidden IE search-focus fixture was active and produced one non-reproducible 2019 search failure. An isolated 2019 rerun passed 34/34. The harness now makes only that audit input read-only after seeding its query; the clean v3 full run passed 136/136 installed-runtime scenarios.
- ProgramData installation and installer packaging were not requested and were not run (`Install: False`).

### Needs Revit Check

- Execute Next Work Queue item 28. The remaining uncertainty is Autodesk runtime behavior: confirm parent placement keeps nested instances, direct/copy-pasted child placement rolls back before commit, dual-use family placement remains allowed, and Admin ON bypass works in both central and differently named local files.

## 2026-07-15 Composite System Type Detailed Components

### Finding

- `StairsType` and `RailingType` were already included in the supported System Type catalog, but the v3 fingerprint contained only identity, classification, segment/material/shape, routing preferences, and compound layers.
- Stair Run/Landing/Support references and railing Handrail/Top Rail/Rail Structure/Baluster configuration were therefore absent from both the fingerprint and detached detail view.
- Autodesk API metadata inspection confirmed the required properties and indexed component collections are available across the installed Revit 2019/2023/2025/2027 APIs. Revit 2021 uses the shared 2019-2023 host build.

### Fix

- Backup: `_backups\system-type-detailed-components-20260715-185809` (95 source, audit, and shared UI files).
- Added shared `SystemTypeDetailedComponentSnapshotService`. A precise scan now walks public `StairsType` and `RailingType` configuration references, including Stair Run/Landing/Support types, Handrail/Top Rail, rail structure, Baluster placement/pattern/post collections, and referenced loadable family/type content fingerprints.
- Referenced System `ElementType` definitions are recursively captured with stable paths and scalar values. Referenced `FamilySymbol` rows include category, family, type, and the deep loadable-family fingerprint.
- System fingerprint schema now uses `SYSFP|v4` with an `S10` detailed-component signature when a supported item was captured by a precise scan. Other System Types and disabled comparisons retain the exact v3 comparison surface.
- Standard snapshot schema is now `7`. Older Railing/Stair snapshots are not interpreted as an empty configuration; they are marked as requiring a new precise scan before detailed comparison can pass.
- Added the administrator checkbox `System Type 상세 구성 요소 비교 (Railing, Stair 등)`. It defaults to enabled, keeps the shared policy as its baseline, persists the explicit PC/user choice separately, refreshes the current browser immediately, and participates in comparison/cache keys.
- Enabled mode compares component identity/value/content fingerprints, promotes a component mismatch to `DifferentFromStandard`, and appends structured `Detailed Components` and `Component Differences` tables to both main and detached System detail views.
- Disabled mode computes the prior v3 comparison fingerprint but keeps the captured component/configuration source in the saved comparison report. Optional Railing/Stair sections are hidden only at final row/detail rendering boundaries, so changing the preference later can reveal the retained data without another model check; mandatory Curtain Panel data is never hidden.
- The shared measurement-unit renderer previously redefined `renderSystemDetailTable` after the component extension and silently removed the new tables. The component extension now wraps the final unit-aware renderer, so routing/layer units and Railing/Stair component tables coexist.

### Automated Coverage

- Contract/static QA covers the new action prefix, policy storage/default, v4 signature, precise-only certification, snapshot persistence, detailed difference generation, final render policy guard, main/detached tables, and the Railing audit fixture across all three host source trees.
- The IE `WebBrowser` harness selects an `AUDIT_GUARDRAIL` row and verifies Top Rail, Handrail, Baluster, the changed Baluster type, two structured tables, and both Korean/English headings.
- A dedicated disabled scenario verifies zero component tables and zero Railing component reference strings in Korean and English for every installed runtime target.
- Static/action/contract checks: PASS for all source hosts (`78` generated actions, `253` exact routes, `63` prefix routes, `11` browser functions each).
- Revit 2019/2021/2023/2025/2027 Release build and Stage verification: PASS with zero compile errors. Revit 2021 runtime remains expected `SKIP runtime-not-installed` while its shared host build and Stage payload pass.
- Full quality gate: `145` scenarios, `144 OK`, `0 FAIL`, `1 SKIP`; `25` warnings are expected handled alerts from intentionally unavailable actions. Report: `artifacts\family-browser-ui-audit\20260715-system-components-final\quality-gate-summary.md`.
- 2,000-row performance/cache gate: PASS. Shell `3-13ms`, cold usable `422-821ms`, warm usable `278-424ms`, and 1,000-row filter `10-13ms` on installed runtime targets.
- Stage DLL SHA256: 2019/2021/2023 `A0DB6CF8E84098588F638AC7C362FFD5F7B578E61D0D011B1B43AE94BF452776`; 2025 `2DD648E96BDE8A8EAC6D44D73C7F24310881465AD50A6FA136C91B602C52EA48`; 2027 `169FCD810FEEB0B8FEFC5F640EA528FF2F5131683A03FB9B7AA47904A0DD13F3`.
- ProgramData installation and installer packaging were not requested and were not run (`Install: False`).

### Needs Revit Check

- Run a new precise standard scan containing real Railing and Stair types. Confirm the detached detail tables match Revit's selected Run/Landing/Support, Handrail/Top Rail, rail structure, Baluster pattern/post, and referenced family/type definitions.
- Compare a project with one changed nested Railing/Stair component while the checkbox is enabled, then disabled. Enabled must classify the parent System Type as different and show the exact row; disabled must ignore and hide only this detailed component layer while preserving routing/layer/parameter comparisons.
- Revit 2021 runtime is not installed on this PC, so only its shared 2019-2023 binary and Stage package were verified automatically.

## 2026-07-16 Railing, Stair, And Curtain Panel Persisted Units

### Problem And Scope

- Railing/Stair detailed components and mandatory Curtain Panel dependencies were fingerprinted and rendered, but their referenced FamilySymbol type parameters were mostly flattened display strings. Length values therefore could not use the same persisted `mm/in` display control as Routing Preferences and compound layers.
- Curtain Panel needed the same behavior for both a curtain host's referenced/default panel and a direct `PanelType` row. Optional Railing/Stair comparison policy must not disable the mandatory Curtain Panel surface.

### Fix

- Backup: `_backups\system-component-unit-switch-20260716-080118` (26 source, fixture, harness, contract, and audit files).
- Shared detailed-component capture now records referenced FamilySymbol type parameters with a stable value kind. Length values preserve raw Revit internal feet using `G17`; angles and Yes/No values retain their semantic kinds instead of being guessed from rendered text.
- Standard snapshot schema is now `9`. Main and detached System details receive paired structured `@component` / `@component-diff` records while retaining legacy `@row` records for migration compatibility.
- Added shared IE-compatible `FamilyBrowserSystemDetailedComponentUnitUi`. Railing, Stair, Curtain Panel dependency, and direct PanelType tables containing a Length row now expose the same `mm/in` selector used by Routing Preferences and compound layers.
- All visible selectors are synchronized, and the preference remains shared at `%LOCALAPPDATA%\KKY\FamilyBrowser\Settings\measurement-unit.txt`. First use defaults to `mm`; closing and reopening after selecting `in` restores `in` without rebuilding the dashboard HTML.
- Legacy snapshots remain readable as text. They are not converted from rounded text because that could produce incorrect units; a new precise scan activates structured conversion.
- The administrator `System Type 상세 구성 요소 비교 (Railing, Stair 등)` option still controls only Railing/Stair data. Curtain Panel capture, comparison, tables, and unit switching remain mandatory when that option is disabled.

### Automated Coverage

- Railing fixture checks `914.4 mm <-> 36 in` plus a changed Baluster offset `76.2/152.4 mm <-> 3/6 in`, two synchronized selectors, and persisted selection.
- Curtain Wall dependency checks `1219.2/1524 mm <-> 48/60 in`; direct PanelType checks `76.2/152.4 mm <-> 3/6 in`. Enabled and disabled optional-component scenarios both verify mandatory Curtain Panel tables.
- Static/action/contract QA requires schema 9, structured component and difference rows, the shared unit renderer, FamilySymbol parameter capture, and Curtain Panel width/thickness fixtures across all three host trees.
- The first full run received external keyboard text while the hidden IE search-focus fixture was active. The audit input now becomes read-only before focus and reseeds its query after focus. An isolated 2019 rerun passed `36/36`, and the clean five-version rerun passed with zero failures.

### Verification And Packaging

- Final quality gate: static/contract and nested-family propagation PASS; Revit 2019/2021/2023/2025/2027 Release build, Stage generation, Stage verification, and 2,000-row performance/cache PASS.
- Full IE `WebBrowser` HTML/click/language/layout/detail harness: `145` scenarios, `144 OK`, `0 FAIL`, `1 SKIP`. The skip is expected Revit 2021 `runtime-not-installed`; its shared host build and Stage payload passed.
- Final report: `artifacts\family-browser-ui-audit\20260716-system-component-units-final-v2\quality-gate-summary.md`.
- Stage DLL SHA256: 2019/2021/2023 `023CCB6364774ED704D4AC2255673FE76B3761E02521D9D9DAEAF9F459F2A013`; 2025 `D4AB4B62CB4877C9A50240DA82C9FEBC35634B897885B1E728D27C50E87C2841`; 2027 `1F4BAE0083FC8268248A50F72B6A71B6462090E4D5020E62E25A99898F6D9E6E`.
- Preserved existing installers and bypassed date-based cleanup. New installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_component-units-20260716_Setup.exe`.
- Desktop copy: `C:\Users\kkyki\Desktop\KKY_FamilyBrowser_v1.0_2019_2021_2023_2025_2027_20260716_component_units_Setup.exe`; size `3,575,558` bytes (`3.41 MB`), PE `MZ` header PASS, SHA256 `8F6D9365438BB14E8EBBB0CE1257A24A96B0B652E0EB7C3DD8F090470717AF80`.
- The installer was packaged but not executed. ProgramData was not changed.

### Needs Revit Check

- Execute Next Work Queue item 30. A new precise standard scan is required because schema 9 raw value metadata cannot be reconstructed safely from older snapshots.
- Verify real Autodesk Railing/Stair component APIs and System/loadable Curtain Panel parameters against Revit's own values in both `mm` and `in`. Revit 2021 remains build/Stage-only until its runtime is installed.

## 2026-07-16 Latest Five-Version Installer

### Build And Packaging

- Audit backup: `_backups\installer-latest-20260716-075421\FAMILY_BROWSER_BUTTON_AUDIT.md`.
- Preserved every existing installer and mail package. The packaging script's date-based cleanup routine was intentionally bypassed.
- Re-ran UI static/action/contract checks against all three source host trees: PASS (`78` generated actions, `253` exact routes, `63` prefix routes, and `11` browser functions per host).
- Rebuilt Release and regenerated/verified Stage for Revit `2019`, `2021`, `2023`, `2025`, and `2027`: PASS, zero compile errors. The existing `NETSDK1137` advisory remains for the 2025/2027 Windows Desktop SDK projects.
- Inno Setup 6.7.3 compiled all five Stage payloads into:
  - `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260716_Setup.exe`
  - Desktop copy: `C:\Users\kkyki\Desktop\KKY_FamilyBrowser_v1.0_2019_2021_2023_2025_2027_20260716_Setup.exe`
- Installer size: `3,570,239` bytes (`3.40 MB`). PE `MZ` header check passed and the Desktop copy hash exactly matches the artifact.
- Installer SHA256: `FA50EA147C97D2FB4B38D1B213396A0EA5BB1DDF222248BD0B3EF3A5899764B6`.
- Stage DLL SHA256: 2019/2021/2023 `8E025F231FACB40FFAD6388FDE6199D2F48C442C0123369EE8DAE2F1FDB00083`; 2025 `7B23D7258DD81D7C2D7B8326E46ED02CECD55729EDF098D0313DAC036C0D84FE`; 2027 `64696FBB9E5992DC7EB55117E332154FF1CD5D37C97884D4D62815A459873C59`.
- No mail-sized ZIP was requested or created.
- Revit 2019 was running during packaging, so the installer was not executed and ProgramData was not changed. Close every Revit process before running the installer as administrator.

## 2026-07-15 Mandatory Curtain Panel Dependency Comparison

### Finding

- Curtain Panels were not fully included in the precise System Type comparison. `OST_CurtainWallPanels` was intentionally excluded from ordinary loadable-family rows, while the System Type fingerprint did not capture the curtain host's default panel reference.
- Curtain `WallType` rows existed, but `AUTO_PANEL_WALL` / `AUTO_PANEL` and the referenced panel family/type content were absent from their fingerprint. `CurtainSystemType` was absent from the supported System Type catalogs, and `PanelType` was excluded because it inherits `FamilySymbol` even though Curtain Panels were also excluded from ordinary Family rows.
- The new administrator option for Railing/Stair detailed components could not be reused for this case: curtain-panel dependencies are part of the curtain host's required definition and must always be scanned and compared.

### Fix

- Backups: `_backups\curtain-panel-required-comparison-20260715-201905` (41 files) and `_backups\curtain-panel-direct-paneltype-20260715-2050` (32 files).
- Extended the shared `SystemTypeDetailedComponentSnapshotService` with the required `CurtainPanelDependencies` group.
  - A curtain `WallType` is detected from `WallType.Kind == WallKind.Curtain`; `CurtainSystemType` and direct `PanelType` rows are also supported.
  - The default panel is resolved through cross-version `BuiltInParameter.AUTO_PANEL_WALL` / `AUTO_PANEL` references.
  - A referenced `FamilySymbol` records category, family, type, type parameters, ElementId references, recursively referenced family types, and the existing deep loadable-family content fingerprint.
  - A direct `PanelType` records its own content fingerprint, every readable type parameter, and dependent ElementId family/type references even when no Curtain Wall currently uses it as the default panel.
  - Precise capture writes a stable marker even when the default is `None`, so an empty setting is distinguishable from legacy or incomplete scan data.
- Added a mandatory `SYSFP|v5` / `S11` curtain-panel signature. The existing optional Railing/Stair surface remains `SYSFP|v4` / `S10`; disabling `System Type 상세 구성 요소 비교 (Railing, Stair 등)` removes only that optional surface and never removes curtain-panel data.
- Added `CurtainSystemType` and `PanelType` to standard registration, project capture, semantic capture, light lookup, selected review, and change-candidate catalogs in all three host trees. Only required Curtain Panel `FamilySymbol` types bypass the normal System Type `FamilySymbol` exclusion, so they remain absent from the ordinary Family list and appear once in the System Type list.
- `PanelType` is explicitly `ReviewOnly`. It is scanned and compared but is not sent through the generic System Type mutation engine.
- Standard snapshot schema is now `8`. Older snapshots cannot be treated as proof of an empty curtain configuration and require a new precise scan.
- Comparison now distinguishes `curtain-components` and `curtain-component-differences`. A changed default/dependent panel promotes the parent curtain host or direct PanelType to `DifferentFromStandard`; missing precise capture becomes `ManualReview` instead of a false match.
- Main and detached System details render localized structured `커튼패널 의존 구성` / `커튼패널 구성 차이` tables. These tables remain visible when the optional Railing/Stair checkbox is off.

### Automated Coverage

- Added `AUDIT_CURTAIN_WALL` with a standard `AUDIT_SYSTEM_PANEL / Glazed`, dependent `AUDIT_PANEL_SUPPORT / Standard`, and current `AUDIT_SYSTEM_PANEL / Solid` difference.
- Added independent `AUDIT_SYSTEM_PANEL_TYPE` coverage with its own content row and a changed `AUDIT_PANEL_INSERT` dependent family type.
- Static and contract QA require the cross-version curtain parameters, required signature, v5 fingerprint, supported Curtain System catalog, comparison path, renderer, and audit fixture in all host trees.
- The IE `WebBrowser` harness selects both the Curtain Wall and direct PanelType rows in enabled and disabled detailed-component scenarios. It requires exactly two curtain tables per selected row and verifies the default panel, direct type, dependent family types, differences, and Korean/English headings.

### Verification

- Static/action/contract and nested-family propagation: PASS for all three source host trees (`78` generated actions, `253` exact routes, `63` prefix routes, `11` browser functions each).
- Revit 2019/2021/2023/2025/2027 Release build, Stage generation, and Stage manifest/payload verification: PASS with zero compile errors. Existing generated-source and Windows-platform analyzer warnings remain unchanged in scope.
- 2,000-row performance/cache gate: PASS.
- Full IE `WebBrowser` HTML/click/language/layout/detail harness: `145` scenarios, `144 OK`, `0 FAIL`, `1 SKIP`. The skip is the expected Revit 2021 runtime-not-installed case; its shared host build and Stage payload passed. The `25` warnings are expected handled alerts from intentionally unavailable actions.
- Final report: `artifacts\family-browser-ui-audit\20260715-curtain-panel-direct-final-v2\quality-gate-summary.md`.
- Stage DLL SHA256: 2019/2021/2023 `8E025F231FACB40FFAD6388FDE6199D2F48C442C0123369EE8DAE2F1FDB00083`; 2025 `7B23D7258DD81D7C2D7B8326E46ED02CECD55729EDF098D0313DAC036C0D84FE`; 2027 `64696FBB9E5992DC7EB55117E332154FF1CD5D37C97884D4D62815A459873C59`.
- ProgramData installation and installer packaging were not requested and were not run (`Install: False`).

### Needs Revit Check

- Execute Next Work Queue item 29. A new precise standard scan is mandatory because schema 8 and the v5 required curtain signature cannot be reconstructed safely from an older snapshot.
- Confirm real Autodesk behavior for both System Panel and loadable Curtain Panel defaults, then alter the default panel and one referenced panel family/type value in the current project. The parent curtain host and independently changed PanelType must be classified as different and show the exact dependency rows even with the Railing/Stair option disabled.
- Revit 2021 runtime is not installed on this PC, so only its shared 2019-2023 binary and Stage package were verified automatically.

## 2026-07-16 Fresh Five-Version Installer 09:07

### Build And Packaging

- Audit backup: `_backups\installer-latest-20260716-090707\FAMILY_BROWSER_BUTTON_AUDIT.md`.
- Preserved all existing installers and mail packages; the packaging script's date-based cleanup routine was bypassed.
- UI static/action/contract checks PASS for all three source host trees (`78` generated actions, `253` exact routes, `63` prefix routes, `11` browser functions each). Nested-family difference propagation also PASS.
- Rebuilt Release, regenerated Stage, and verified payloads for Revit `2019`, `2021`, `2023`, `2025`, and `2027`: PASS with zero compile errors. Existing `NETSDK1137` advisories remain for 2025/2027.
- Installer artifact: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260716-090707_Setup.exe`.
- Installer size: `3,575,558` bytes (`3.41 MB`); PE `MZ` header PASS; SHA256 `5966B80C439162AE2F2718D2B1110A3EECEFCC635303F2638617B5BAFFDB9F2E`.
- The temporary Desktop copy was removed after the user clarified that deliverables must remain in the normal artifact folders.
- Revit was not running. The installer was packaged but not executed, so ProgramData was not changed.
- Mail package backup: `_backups\mail-package-20260716-091014\FAMILY_BROWSER_BUTTON_AUDIT.md`.
- Mail package: `artifacts\family-browser\mail-packages\20260716_01.zip`; size `16,777,451` bytes (`16.00 MB`); SHA256 `1722F7BB9B6C4F3235195958BA8F53C8BEBB2FAA671C5BB819469F376A0AA684`.
- ZIP entries are `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`. The `Setup.exe` SHA256 inside the ZIP matches the installer artifact exactly.

## 2026-07-16 Detailed System Type Option Layout

### Problem And Cause

- The `System Type 상세 구성 요소 비교 (Railing, Stair 등)` control used a flex row with `gap: 10px`. The legacy IE `WebBrowser` does not support flex gap, so the checkbox touched the long title and the setting read like one unstructured text line.

### Fix

- Backup: `_backups\system-detail-option-ui-20260716-0920` (8 source, CSS, harness, contract, and audit files).
- Replaced the gap-dependent layout with IE-safe table cells: a fixed checkbox cell, a flexible title/description cell, and a separated `사용/미사용` (`Enabled/Disabled`) state cell.
- Increased the checkbox to `20px`, added a reliable `14px` cell padding, separated title and help copy vertically, and added a restrained left accent, focus outline, divider, and state dot without changing the route or checkbox semantics.
- The entire row remains clickable and keeps `role=checkbox` plus `aria-checked`; only presentation changed.

### Automated And Visual Verification

- Static/action/contract checks PASS for all three host trees (`78` generated actions, `253` exact routes, `63` prefix routes, `11` browser functions each).
- Revit 2019/2021/2023/2025/2027 Release build, Stage generation, and Stage payload verification PASS with zero compile errors.
- The IE layout audit now requires at least `10px` between the checkbox and copy, vertical title/help separation, a non-overlapping state cell, and no option-edge overflow.
- Focused Revit 2025 IE harness: `36/36 OK`, zero failures in Korean/English and all dashboard/layout scenarios. Report: `artifacts\family-browser-ui-audit\20260716-system-detail-option-ui-2025\ui-harness-summary.json`.
- Visual capture: `artifacts\family-browser-ui-audit\20260716-system-detail-option-ui-2025\Rvt2025-admin-standard-settings-layout-ko-light\preview.png`.
- Stage DLL SHA256: 2019/2021/2023 `ECC0958B9E698637C130EEE0AFA1E3BA5D94D385F7E8469D69FA255F65760931`; 2025 `EB7381246497B000D6769CE6C14E3C4331F6C27353CB2483F288AB0FE41EF4A7`; 2027 `F3B598DABCE34611E9B791FC1896370F8389B268E2111C30071EC0D7692878B4`.
- ProgramData was not changed. The 09:07 installer and `20260716_01.zip` predate this visual fix and were not regenerated in this task.

## 2026-07-16 Post-Option-Layout Installer And Mail Package 09:40

### Build And Packaging

- Audit backup: `_backups\installer-mail-20260716-094035\FAMILY_BROWSER_BUTTON_AUDIT.md`.
- Preserved all existing installer and mail-package artifacts; no Desktop copy was created.
- UI static/action/contract checks PASS for all three host trees (`78` generated actions, `253` exact routes, `63` prefix routes, `11` browser functions each).
- Revit 2019/2021/2023/2025/2027 Release build, Stage regeneration, and Stage payload verification PASS with zero compile errors. Existing `NETSDK1137` advisories remain for 2025/2027.
- Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260716-094007_Setup.exe`; size `3,578,932` bytes (`3.41 MB`); PE `MZ` header PASS; SHA256 `FB440F80FF166DEA0CC67BFCB1CD5F2E8613B66710A78522790F405483719956`.
- Mail package: `artifacts\family-browser\mail-packages\20260716_02.zip`; size `16,777,451` bytes (`16.00 MB`); SHA256 `2162F95C20063B85DE566AF5B5F9030BDB5265958501ADDD58DFD6BBFE96F6CD`.
- ZIP entries are `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`; the embedded installer SHA256 matches the installer artifact.
- The package includes the IE-safe Detailed System Type option layout. The installer was not executed, so ProgramData was not changed.

## 2026-07-16 Family/System Trade Selector And Left Tree Synchronization

### Problem And Cause

- In the Family and System Type lists, the prepared-trade buttons below the search field changed the visible row dataset correctly, but the left `전체 공종` / current-trade tree could retain the previously selected trade name and count.
- The host route updated `_browseDisciplineKey`, while both left trees were still generated from the full cached multi-trade row collection. Their browser-side state also started from the selected trade instead of the root and could preserve a stale category/trade refinement across the host refresh.
- The trade selector visually resembled the ordinary row-status filters, so its role as the standard data target was not sufficiently clear.

### Fix

- Backup: `_backups\browser-trade-tree-sync-20260716-095731`.
- Family and System left trees now derive their root count, trade node, and category nodes only from rows matching the current `_browseDisciplineKey` or its configured slot label. Switching Mechanical to Architectural therefore regenerates the tree as `전체 공종 / 건축`, not the previous `설비` tree.
- Family/System tree state now starts at `All`. `beginBrowseDisciplineSwitch(...)` clears both old tree refinements before the host navigation, updates the selected trade chip immediately, and preserves each chip's ready/unavailable state while rows change.
- The search-area trade selector now has a separate blue-tinted band, stronger label, registered/unregistered visual states, a solid blue selected target, and a compact readiness badge. The styling is static/IE-compatible and wraps without clipping.

### Automated And Visual Verification

- Added audit scenario input `BrowseDisciplineKey`; `viewport-1280-family` now starts from `Architectural` even though the fixture also contains Mechanical data.
- Added IE checks that require one selected chip, an `All Trades` root, exactly one matching current-trade tree root, matching root/trade/row counts, and reset of stale Family/System tree refinements on an immediate trade switch.
- Static/action/contract checks PASS for all three host trees (`78` generated actions, `253` exact routes, `63` prefix routes, `11` browser functions each).
- Revit 2019/2021/2023/2025/2027 Release build, Stage generation, and Stage payload verification PASS with zero compile errors.
- Full IE harness on all three unique host binaries: 2025 `36/36 OK`; shared 2019/2021/2023 host represented by 2019 `36/36 OK`; 2027 `36/36 OK`; total `108 OK`, `0 FAIL`.
- Reports: `artifacts\family-browser-ui-audit\20260716-trade-tree-sync-2025-full` and `artifacts\family-browser-ui-audit\20260716-trade-tree-sync-other-hosts`.
- Visual capture: `artifacts\family-browser-ui-audit\20260716-trade-tree-sync-2025-full\Rvt2025-viewport-1280-family-ko-light\preview.png`; it shows selected `건축`, left `전체 공종 1`, and left `건축 1` with no stale Mechanical label.
- Stage DLL SHA256: 2019/2021/2023 `9C979CA9FC63257B5F32133438DAF0F66FC9A897D34C6BC0E488FDB4F93FFC63`; 2025 `439A52F443162B86ED7957D38C68E2E4417CF0310BB73421A7DF4230123F7AC3`; 2027 `ABC4C81F4D637E605C8C2F32079A096B18B38E8558FEE5EC9222F92DA9530789`.
- ProgramData, installer, and mail package were not changed. The 09:40 installer and `20260716_02.zip` predate this fix.

### Needs Revit Check

- With real registered Architectural and Mechanical standards, switch both directions in the Family and System Type tabs. Confirm the rows, `전체 공종` count, current-trade label/count, categories, and selected chip all change together without requiring Refresh.

## 2026-07-16 Isolated Family/System Column Resizing

### Problem And Cause

- Dragging one Family/System list header changed neighboring column widths slightly even though their stored width values were untouched.
- The resize runtime assigned pixel widths to every `col`, but the list tables still inherited `width: 100%` and an `1180px` important minimum width. IE fixed-table layout redistributed the remaining table width across columns whenever the requested sum and the forced table width differed.
- On first use without saved widths, the runtime also waited until the first mouse-down to freeze the measured baseline, leaving the initial table eligible for proportional reflow.

### Fix

- Backup: `_backups\column-resize-isolation-20260716-102855`.
- `family-browser-row-window.js` now measures and freezes every header/body column as an explicit pixel baseline when resizers are attached, before any drag begins.
- The header and body tables receive `kkyfb-column-width-locked`; their rendered width is always the exact sum of the explicit columns. The shared CSS removes the legacy minimum-width constraint only for these locked list tables, so a drag changes the target column and the table's total width without redistributing space to untouched columns.
- In-memory and persisted width arrays are cloned before storage so a live drag cannot mutate another caller's baseline by reference.
- The IE performance harness now shrinks a real column, compares all untouched header/body `col` widths and rendered `th`/`td` pixel widths, and verifies both tables remain locked to the exact summed width.

### Verification

- UI static/action/contract checks: PASS for all three source hosts (`78` generated actions, `253` exact routes, `63` prefix routes, `11` browser functions each).
- Focused Revit 2025 2,000-row performance/resize gate: PASS for both Family and System tables. Report: `artifacts\family-browser-ui-audit\20260716-column-resize-isolation-focused`.
- Revit 2019/2021/2023/2025/2027 Release build and Stage payload verification: PASS with zero compile errors. Existing advisory/analyzer warnings remain unchanged in scope.
- Full IE harness: Revit 2019 `36/36 OK`, 2023 `36/36 OK`, 2025 `36/36 OK`, and 2027 `36/36 OK`; Revit 2021 runtime `SKIP runtime-not-installed`. No UI failures were reported. The initial combined command exceeded the 10-minute shell limit after completing 2019/2023/2025 and part of 2027, so 2027 was rerun independently to a complete passing summary.
- Reports: `artifacts\family-browser-ui-audit\20260716-column-resize-isolation-full`, `artifacts\family-browser-ui-audit\20260716-column-resize-isolation-2027`, and `artifacts\family-browser-ui-audit\20260716-column-resize-isolation-2021-skip`.
- Stage DLL SHA256: 2019/2021/2023 `8799939FC4D921CF900C086E13BF95FB2CB377E4F6E4F6C76AE9E8BD0B1D4F7B`; 2025 `F3109EB2C4E8B4903391AC69677EB28E710D2E9F4CE8DF67E091EE2F7311748F`; 2027 `B3D73A6F9FCA1A501998B46AB0280FDD0C4FE9F3551D4D6D6AE9270F7028EC23`.
- ProgramData, installer, and mail package were not changed.

### Needs Revit Check

- In a real Family and System Type list, drag one narrow and one wide column in both directions. Confirm only that column changes, horizontal scrolling grows or shrinks with the total table width, header/body remain aligned, and the same widths return after reopening the browser.

## 2026-07-16 Column Resize Installer And Mail Package 10:58

### Build And Packaging

- Backup: `_backups\installer-mail-column-resize-20260716-105817`.
- Added `-PreserveExistingArtifacts` to `Build-FamilyBrowserInstaller.ps1` and used it for this build, so no previous installer or mail package was removed.
- Rebuilt Release, regenerated Stage, and verified payloads for Revit `2019`, `2021`, `2023`, `2025`, and `2027`: PASS with zero compile errors. Existing `NETSDK1137` advisories remain for 2025/2027.
- Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260716-105843_Setup.exe`; size `3,577,020` bytes (`3.41 MB`); PE `MZ` header PASS; SHA256 `08B469C68E0AA5F4A6874FBC4324AC25E5A457149EDEA223ADFEC23FF8424713`.
- Mail package: `artifacts\family-browser\mail-packages\20260716_03.zip`; size `16,777,291` bytes (`16.0001 MB`); SHA256 `21AC04A781323EF64FE35B102B303E53EA92CDF2723455A289143719BB2E5C0D`.
- ZIP entries are `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`. The embedded `Setup.exe` SHA256 exactly matches the standalone installer.
- Updated `latest-build.json`, `latest-mail-package.json`, and both SHA256 text files to the new artifacts.
- Revit was running, so the installer was not executed and ProgramData was not changed.

## 2026-07-16 Current Trade Label Final Synchronization

### Problem And Cause

- Real Revit testing confirmed a narrower regression after the earlier trade-tree fix: switching the prepared trade changed the active chip, Family/System rows, categories, tree key, and counts correctly, but the visible level-1 trade name below `전체 공종` could remain on the previous trade.
- The earlier fix scoped the tree rows to `_browseDisciplineKey`, but the rendered level-1 name still came from the first cached row's `DisciplineLabel`. The immediate browser feedback also updated the chip and filters without rewriting that existing tree text node. This allowed a stale cached display label to survive even while the actual target data had already changed.

### Fix

- Backup: `_backups\browse-trade-tree-label-sync-20260716-111648`.
- Added `ResolveCurrentBrowseTreeLabel(...)` to all three host source trees. When an explicit browse target is selected, both Family and System level-1 tree names now resolve from `ResolveBrowseSlot(_standardPolicy)` / `ResolveSlotLabel(...)`; cached row text is fallback only.
- Added IE-compatible `syncBrowseTreeTradeLabel(...)`. `beginBrowseDisciplineSwitch(...)` now copies the clicked trade chip's direct text into the existing Family/System level-1 tree node immediately, before resetting tree filters and before the host refresh completes.
- The implementation changes only the visible label. Existing tree keys, row filtering, categories, counts, paging, selection, and isolated column widths remain unchanged.

### Automated Verification

- Expanded the IE harness to compare the active prepared-trade chip's direct label with the current level-1 left-tree label on initial render and after an inactive trade is switched. The harness restores the original label after the interaction test.
- Expanded static QA to require the selected-slot label resolver in both Family and System tree builders and the immediate browser-side label synchronization call.
- Static/action/contract checks: PASS for all three source hosts (`78` generated actions, `253` exact routes, `63` prefix routes, `11` browser functions each).
- Revit 2019/2021/2023/2025/2027 Release build and Stage generation: PASS with zero compile errors. Stage manifest/payload verification: PASS for all five targets.
- Shared 2019/2021/2023 host: the Revit 2019 full run produced `36` result JSON files with zero failures; the interrupted combined run also completed `24` Revit 2023 scenarios with zero failures before the outer five-minute command limit. Revit 2021 runtime remains `SKIP runtime-not-installed`, while its shared binary and Stage payload passed.
- Revit 2025 full IE harness: `36/36 OK`, zero failures. Report: `artifacts\family-browser-ui-audit\trade-tree-label-20260716-2025`.
- Revit 2027 full IE harness: `36/36 OK`, zero failures. Report: `artifacts\family-browser-ui-audit\trade-tree-label-20260716-2027`.
- Stage DLL SHA256: 2019/2021/2023 `664F20008B5860BD436F760705F26E4DA077D1E9101177EBDBDA541CA0F8B92B`; 2025 `EB8A61968DCDDB032EE0F21A8DC701F089ACB3DD776AF64AAE875FBF22DC4C3D`; 2027 `7BEB8D7D8D29A32FD44A679813B3704C66348E8FC17AF126A1417EF6CA135CDE`.
- Revit was running, so ProgramData installation was not attempted. The 10:58 installer and `20260716_03.zip` predate this label fix and were not regenerated.

### Needs Revit Check

- After installing a package that includes this source revision, switch Architectural and Mechanical in both directions on the Family and System Type tabs. The prepared-trade chip, rows, `전체 공종` count, current-trade name/count, and categories must all change together without Refresh.

## 2026-07-16 Missing Management Folder Onboarding Recovery

### Problem And Cause

- A PC that could not reach the management-folder path published by the homepage showed only the general `미확인` readiness state. The intended startup guidance and TEST-folder setup choice did not reliably appear, leaving a first-time user with no recovery action.
- The missing-folder state was calculated after the startup shell's first `DocumentCompleted` event had already fired. A later shell refresh does not always produce another full document-completed event, so the pending onboarding dialog could remain queued forever.
- The old prompt was also restricted to the Home tab and suppressed by a process-wide one-shot flag. Closing it left no persistent recovery control in the browser.

### Fix

- Backup: `_backups/managed-folder-onboarding-recovery-20260716-115017`.
- All three host source trees now keep the missing-folder state per browser form and queue the onboarding immediately after startup preload completion as well as after `DocumentCompleted`. The prompt is no longer dependent on which tab was active when the browser opened.
- When the homepage path is unreachable, a structured `관리폴더 연결 필요` window now explains two paths clearly: connect the company network/VPN and contact the administrator or deployment owner for normal use, or configure a TEST management folder only for an isolated test environment.
- `TEST 관리폴더 설정` opens folder selection and accepts only an accessible UNC path or mapped internal network drive. Local disks and personal folders remain rejected. The chosen override is stored in the existing per-user `managed-folder-override.txt`, and generated managed files carry a warning against manual edit, move, rename, or deletion.
- Closing the startup prompt no longer leaves the user stranded. Home now keeps a full-width recovery banner with `홈페이지 경로 다시 확인` and `TEST 관리폴더 설정`. A successful homepage retry or TEST setup removes the banner immediately.
- Added exact action routes `managed-folder-retry` and `managed-folder-test-setup`, Korean/English contract coverage, and an unavailable-management-folder audit scenario. The banner is also asserted to be absent when a usable path exists.
- The UI audit screenshot runner no longer treats Edge's successful `bytes written` stderr message as a terminating failure; PNG existence and size remain the actual screenshot success check.

### Automated And Visual Verification

- UI static and action/contract checks: PASS for all three source hosts. Contract counts after this change: `80` generated actions, `255` exact routes, `63` prefix routes, and `11` browser functions per host.
- Full IE `WebBrowser` harness: Revit 2025 `38/38 OK`; shared 2019/2021/2023 host represented by 2019 plus Revit 2027 `76/76 OK`; total `114 OK`, `0 FAIL`.
- The six Korean/English missing-management-folder scenarios across the three unique host binaries each expose exactly one retry action and one TEST setup action; all `25` visible click candidates route successfully with zero failures.
- Reports: `artifacts/family-browser-ui-audit/managed-folder-onboarding-20260716-2025` and `artifacts/family-browser-ui-audit/managed-folder-onboarding-20260716-final3`.
- Visual capture: `artifacts/family-browser-ui-audit/managed-folder-onboarding-20260716-2025/Rvt2025-admin-home-managed-folder-unavailable-ko-light/preview.png`. The recovery banner is fully visible below the header, does not overlap the sidebar or dashboard, and keeps both actions readable.
- Revit 2019/2021/2023/2025/2027 Release build and Stage generation completed with zero compile errors. Stage manifest/payload verification: PASS for all five targets.
- Stage DLL SHA256: 2019/2021/2023 `BB45190C95C30FA3062AAACE18EFFAF94CEDFDF6C6310264F1D897D2ADA3F81F`; 2025 `5B8B014199ADF5F15EA7DE93C99DF195BFE60DBFA95C6A3CC030F2E8385113C3`; 2027 `BF7D0380199034477E8C72E9336FD184C0ED0AAD21D1FD372CF24CA29B2A5324`.
- ProgramData, installer, and mail package were not changed in this task.

### Needs Revit Check

- On a PC where the homepage-published path is genuinely unreachable, open Family Browser and confirm the structured guidance window appears once per browser opening regardless of the initial tab.
- Close the window and confirm the Home recovery banner remains. Connect VPN or the company network, choose `홈페이지 경로 다시 확인`, and confirm the banner disappears when the published path becomes reachable.
- With the homepage path still unavailable, choose an internal UNC or mapped network folder through `TEST 관리폴더 설정`, restart Revit, and confirm the persisted TEST management folder is reused. Confirm a local `C:` folder is rejected and no existing managed files are modified manually.

## Standing Rule: ProgramData Is Part Of Every Completed Source Change

- Effective 2026-07-16, every completed Family Browser source change must finish with Release build, Stage verification, ProgramData deployment, installed-payload verification, and Stage-versus-installed SHA256 comparison for Revit `2019/2021/2023/2025/2027`.
- Installer and mail-package generation remain separate requested deliverables, but ProgramData is no longer optional after a source modification.
- If Revit is running or the installed DLL is locked, do not report the change as fully deployed. Record the pending ProgramData step, close or ask to close Revit, then complete deployment and verification as soon as the lock is released.

## 2026-07-16 Management Folder Onboarding ProgramData Deployment 12:45

- Revit process count before deployment: `0`.
- Existing installed Family Browser payloads were backed up to `_backups/programdata-before-managed-folder-onboarding-20260716-124515`.
- The first non-elevated copy was correctly rejected by Windows ACLs. Deployment was rerun through an elevated installation process and completed successfully.
- ProgramData install verification: `Installed OK` for Revit `2019`, `2021`, `2023`, `2025`, and `2027`.
- Stage-versus-installed DLL SHA256 comparison: all five `Match=True`.
- Installed DLL SHA256: 2019/2021/2023 `BB45190C95C30FA3062AAACE18EFFAF94CEDFDF6C6310264F1D897D2ADA3F81F`; 2025 `5B8B014199ADF5F15EA7DE93C99DF195BFE60DBFA95C6A3CC030F2E8385113C3`; 2027 `BF7D0380199034477E8C72E9336FD184C0ED0AAD21D1FD372CF24CA29B2A5324`.
- Deployment log: `artifacts/family-browser/programdata-install-20260716-1245.log`.

## 2026-07-16 Generated Family Browser Ribbon Icon 13:01

### Asset Change

- Replaced the temporary blue-window ribbon mark with a Family Browser icon generated in the same bright 3D visual family as the actual KKY Tool Revit icon: pale-blue rounded plate, BIM sheets, blue family folder with an `F`, family components, and a search lens.
- The generated source's uniform black outer canvas was converted to alpha while preserving the dark-blue icon outline. Source alpha audit: `616,697/1,572,516` fully transparent pixels and `14,744` antialiased partial-alpha pixels.
- New project assets: `KKY_FamilyBrowser_SharedUi/family-browser-ribbon-source.png`, `family-browser-ribbon-32.png`, and `family-browser-ribbon-16.png`.
- The Revit assets are true `32x32` and `16x16` RGBA PNGs. Both have `cornerA=0`; visible alpha bounds are `(3,2)-(29,29)` at 32px and `(1,1)-(15,15)` at 16px.
- Asset SHA256: source `EF9975FCF625C0E9F6A16F9237145DBA44E0C0B7656552E53E84DABD660761F6`; 32px `3EA9C665A9E71F082B176A4D3E684037747FC12A4228DCA75A34F01F68CE9AD0`; 16px `FD1E8ECC6EA9E4641CBDD17CC0198A00A92DA4C0A21D1740C635C908B769D153`.
- Previous ribbon assets backup: `_backups/family-browser-ribbon-icon-before-generated-20260716-125742`.

### Build And Deployment Verification

- Revit 2019/2021/2023/2025/2027 Release build completed with zero compile errors. Existing analyzer/advisory warnings remain unchanged in scope.
- Stage manifest/payload verification: PASS for all five targets.
- Extracted embedded-resource hashes from the shared 2019/2021/2023 host, 2025 host, and 2027 host. All six embedded 16px/32px resources exactly match the source PNG hashes.
- ProgramData backup: `_backups/programdata-before-family-browser-ribbon-icon-20260716-130139`.
- ProgramData install verification: `Installed OK` for Revit `2019`, `2021`, `2023`, `2025`, and `2027`.
- Stage-versus-installed DLL SHA256 comparison: all five `Match=True`.
- Installed DLL SHA256: 2019/2021/2023 `893B899CEF1109DACDC39F88627D40A938E482F7DAD15053A84885780ED18DE0`; 2025 `F7C7068AC93B2CE25096BE8ADCC3B616ECE23E1647DF9640BABC3D54A96FF5ED`; 2027 `CB511E0EE877E4C584D4DA6CA9C4B7E03650283698DB62E534B0B42910DDB7AD`.
- Deployment log: `artifacts/family-browser/programdata-install-ribbon-icon-20260716-1301.log`.

### Needs Revit Check

- Restart Revit and inspect the Family Browser command in the `KKY Tools` ribbon at normal and high-DPI scaling. Confirm the transparent corners blend with the ribbon, the folder/search silhouette remains recognizable, and the icon is not visually clipped at 16px or 32px.

## 2026-07-16 Authoritative System Type Installer And Mail Package 14:03

### Build And Packaging

- Backup: `_backups\installer-mail-system-type-authoritative-20260716-140258`.
- Ran `Build-FamilyBrowserInstaller.ps1 -Version 1.0 -Label latest-20260716-140154 -MailPackageMinimumMB 15.5 -PreserveExistingArtifacts`; all previous installer and mail artifacts were preserved.
- Revit 2019/2021/2023/2025/2027 Release compilation completed with zero errors. The existing `NETSDK1137` advisory remains for 2025/2027.
- Stage and installed-payload verification both passed for all five targets. Stage-versus-ProgramData DLL SHA256 comparison is `Match=True` for every target: 2019/2021/2023 `2735A0DF9DED33D0F2C24A58C4343D8A6269749FDF72F1F667F1DCB8091CDFFD`; 2025 `841921BDCB15A0BE6E52AC1B790B242ED9552F96499CFFFD360BBD4CABE08B8A`; 2027 `C94B2682CE9214BA1281F51BB8E77B2D3D8FDE9783E400200A2DC9939C20EDC6`.
- Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260716-140154_Setup.exe`; size `3,590,887` bytes; PE `MZ` header PASS; SHA256 `EF88DBF6301895B4552FC998A19685D085A4867BF378AF9A96657D8C135894DC`.
- Mail package: `artifacts\family-browser\mail-packages\20260716_04.zip`; size `16,777,451` bytes (`16.00 MB`); SHA256 `28261CE9B40BCD37D06E9533DBECE6E9E7160E9BA52912CCD2227CA9EEE8CC07`.
- ZIP entries are `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`. Embedded `Setup.exe` SHA256 exactly matches the standalone installer.
- Updated `latest-build.json`, `latest-mail-package.json`, and the ZIP checksum file to the validated artifacts. No Desktop copy was created.

## 2026-07-16 System Type Detail Preference And Comparison Recovery 15:06

### Problem And Cause

- `System Type 상세 구성 요소 비교 (Railing, Stair 등)` was stored only in the shared `standard-policy.json`. A Current Model Check, homepage policy refresh, or prepared-policy reload could replace that in-memory value, so a checked option appeared unchecked again.
- System Type capture itself already collected Railing/Stair detail. The loss happened later: when the option was off, `ProjectStandardComparisonService.BuildSystemTypeDetailSummary(...)` destructively removed `components` and `component-differences` from the comparison report before it was saved.
- System rows reconstructed from a saved comparison report also omitted `BrowserDetailKey` and `BrowserManifestSourceKey`. Therefore an older stripped report could show only identity fields such as `Class`, and checking the option again had no V2 detail record route to recover the full standard configuration.

### Fix

- Source backup: `_backups\system-detail-preference-and-report-recovery-20260716-143949` (`29` files).
- Added the per-user preference `%LOCALAPPDATA%\KKY\FamilyBrowser\Settings\system-type-detail-components.txt` in all three host trees. An explicit `on` or `off` choice now overrides the shared policy baseline and is used by the dashboard, Revit bridge, compare/stamp commands, Family load, and System Type apply paths. The full browser reset clears this local preference.
- Added an audit-only nullable override so real workstation preferences cannot contaminate the deterministic enabled/disabled UI scenarios.
- Removed destructive filtering from comparison-report creation. The checkbox now controls fingerprint/difference classification and final visibility, while the saved report retains the captured source detail for a later ON/OFF switch.
- `AppendSystemRowsFromComparison(...)` now receives the standard snapshot, matches each comparison row by class/category/type with class/type fallback, and stores the V2 detail/source keys.
- Lazy System detail hydration now preserves an already complete comparison summary, but if a legacy row lacks the optional `components` section it loads the standard V2 detail record on selection. A genuinely old standard snapshot without V2 detail still requires one new precise scan.

### Automated Verification

- Static/action/contract checks: PASS for all three source hosts (`80` generated actions, `255` exact routes, `63` prefix routes, `11` browser functions each). New guards require the local preference service, audit isolation, non-destructive report storage, comparison-row V2 keys, and legacy detail hydration.
- Revit `2019/2021/2023/2025/2027` Release build and Stage verification: PASS with zero compile errors. Existing 2025/2027 analyzer/advisory warnings are unchanged in scope.
- Nested-family difference propagation and authoritative System Type apply source tests: PASS.
- 2,000-row performance/cache gate: PASS.
- Full IE `WebBrowser` gate: Revit 2019, 2023, 2025, and 2027 produced `152 OK`, `0 FAIL`; Revit 2021 runtime is `SKIP runtime-not-installed`, while its shared binary and Stage payload passed. Report: `artifacts\family-browser-ui-audit\system-detail-persistence-20260716\quality-gate-summary.md`.
- Enabled fixture check: `AUDIT_GUARDRAIL` contains one `@section\tcomponents`, four `@component\tcomponents`, and `Rail Height`. Disabled fixture check: the same scenario retains its rows but contains zero optional component section/row tokens and no `Rail Height`.
- Stage DLL SHA256: 2019/2021/2023 `588311839779983D97014179CF0B60681252E636742CE7BD92D2AD9FFB950E56`; 2025 `E375D94F5251235E7B2F5A83CC7331748B98327187042EF1439D9EF41FA27823`; 2027 `3B6D4D79419ABFB118A16A5422BAA6348BE8145C9D54931102DEBDF9E367578F`.

### ProgramData Status

- Revit process count was `0`. Existing installed payload backup: `_backups\programdata-before-system-detail-persistence-20260716-150106`.
- The normal install was denied by the ProgramData ACL. Two visible UAC `RunAs` attempts both returned `사용자가 작업을 취소했습니다`, so ProgramData was not changed.
- Stage-versus-installed SHA256 remains `Match=False` for all five targets. Installed DLLs are still the prior 14:03 package: 2019/2021/2023 `2735A0DF9DED33D0F2C24A58C4343D8A6269749FDF72F1F667F1DCB8091CDFFD`; 2025 `841921BDCB15A0BE6E52AC1B790B242ED9552F96499CFFFD360BBD4CABE08B8A`; 2027 `C94B2682CE9214BA1281F51BB8E77B2D3D8FDE9783E400200A2DC9939C20EDC6`.
- ProgramData deployment and installed-payload verification remain pending until one UAC elevation is approved. No installer or mail package was requested/generated for this change.

### Needs Revit Check

- After elevated deployment, execute Next Work Queue item 31 with one real Railing or Stair type. The automated fixture proves persistence routing and visibility boundaries, but actual Autodesk component capture plus recovery from this PC's existing managed comparison files still needs Revit runtime confirmation.

## 2026-07-16 System Type Detail Preference ProgramData Deployment 15:10

- Revit process count before deployment: `0`.
- The elevated install was approved and completed for Revit `2019`, `2021`, `2023`, `2025`, and `2027`.
- Installed-payload verification: `Installed OK` for all five targets.
- Stage-versus-installed DLL SHA256 comparison: all five `Match=True`.
- Installed DLL SHA256: 2019/2021/2023 `588311839779983D97014179CF0B60681252E636742CE7BD92D2AD9FFB950E56`; 2025 `E375D94F5251235E7B2F5A83CC7331748B98327187042EF1439D9EF41FA27823`; 2027 `3B6D4D79419ABFB118A16A5422BAA6348BE8145C9D54931102DEBDF9E367578F`.
- Deployment log: `artifacts\family-browser\programdata-install-system-detail-20260716.log`.
- Runtime check remains: turn the detailed System Type component comparison option ON, run Current Model Check, return to Standards and System Types, and verify that the option stays ON and a real Railing/Stair detail shows its component tables. Then repeat once with the option OFF to confirm the preference remains OFF and only the optional component sections are hidden.

## 2026-07-16 System Type Detail Preference Installer And Mail Package 15:12

- Backup of the previous latest-artifact pointers: `_backups\installer-mail-system-detail-persistence-20260716-151151`.
- Ran `Build-FamilyBrowserInstaller.ps1 -Version 1.0 -Label latest-20260716-151151 -MailPackageMinimumMB 15.5 -PreserveExistingArtifacts`; all previous installer and mail-package files were preserved.
- Revit `2019`, `2021`, `2023`, `2025`, and `2027` Release builds completed with zero compile errors. The existing `NETSDK1137` advisory remains for 2025/2027.
- Stage and installed-payload verification passed for all five targets. Stage-versus-ProgramData DLL SHA256 comparison is `Match=True` for every target.
- Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260716-151151_Setup.exe`; size `3,586,217` bytes; PE `MZ` header PASS; SHA256 `9848AB18C686B91DD2BA62C16E2DA3063E8CE263A4A1D1477A5D0E2AE38A8017`.
- Mail package: `artifacts\family-browser\mail-packages\20260716_05.zip`; size `16,777,446` bytes (`16.00 MB`); SHA256 `0160A1DC7D8F56A7A7EA35A572501007CD5039EAA189F4E03D5857F410F1F6F0`.
- ZIP entries are `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`. The embedded `Setup.exe` SHA256 exactly matches the standalone installer.
- `latest-build.json`, `latest-mail-package.json`, and their SHA256 pointer files were updated to the validated artifacts. No Desktop copy was created.

## 2026-07-16 TEST Management Folder Homepage Return And Migration 16:12

### Problem And Cause

- A TEST management-folder selection persisted through `managed-folder-override.txt` and disabled homepage bootstrap. Even after a homepage path became available, the saved TEST override won again on the next startup.
- The top Refresh action refreshed homepage security keywords only. While bootstrap was disabled by the TEST override, it did not perform a non-mutating homepage managed-folder availability check.
- Because a valid TEST folder counted as an available management folder, Home had no recovery or transition UI and the user could not intentionally switch or migrate back to the centrally published homepage path.

### Fix

- Source backup: `_backups\managed-folder-homepage-return-20260716-153310`.
- Added a live homepage managed-folder probe to the 2019-2023, 2025, and 2027 bootstrap services. The probe ignores only the locally persisted TEST-disable token, respects an explicit bootstrap URL, bypasses HTTP cache, verifies the published path, and does not change the currently active folder.
- Top Refresh now checks for a homepage folder while a TEST override is active. It keeps the TEST folder active until the user chooses a transition, so refresh cannot silently redirect a working session.
- Home now always identifies an active TEST folder. When the homepage folder is available it offers `홈페이지 경로로 변경`; administrator mode additionally offers `기존 데이터 이관 후 변경`. The same controls are available in Standards administration.
- Added a shared migration service for Config, RevitVersions, StandardLists, Requests, Logs, Diagnostics, OperationLogs, StandardChangeCandidates, and related managed data. It compares file bytes, skips temporary/reparse content, rebases JSON string paths rooted under the TEST folder, and never deletes the TEST source folder.
- Migration performs a preflight before copying. A different existing central managed file is treated as a conflict and blocks activation rather than being overwritten. Log/diagnostic name collisions are skipped and reported.
- Successful activation order is migration when requested, clear the disabled-bootstrap marker, apply and verify the exact homepage policy path, then remove the TEST override pointer. If activation or pointer removal fails, the previous TEST override is restored.
- Migration remains restricted to administrator mode and `ManagePolicy` permission. Switching without migration is available to normal users because it does not write managed data.

### Automated Verification

- Static/action/contract checks: PASS for all three host trees (`80` generated actions, `257` exact routes, `63` prefix routes, and `11` browser functions each).
- The migration fixture verified recursive copy, JSON path rebasing, TEST source retention, conflict blocking, and preservation of an existing different destination policy file.
- Full IE `WebBrowser` gate: Revit 2025 produced `42 OK`, `0 FAIL`; Revit 2019/2023/2027 produced another `126 OK`, `0 FAIL`. Revit 2021 runtime is `SKIP runtime-not-installed`, while its shared binary, contract scenarios, Stage payload, and installed payload passed.
- New Korean and English scenarios cover TEST-active/homepage-unavailable and TEST-active/homepage-available states. Visual result: `artifacts\family-browser-ui-audit\managed-folder-transition-20260716-2025\Rvt2025-admin-home-test-folder-homepage-available-ko-light\preview.png`.
- Audit reports: `artifacts\family-browser-ui-audit\managed-folder-transition-20260716-2025`, `artifacts\family-browser-ui-audit\managed-folder-transition-20260716-remaining`, and `artifacts\family-browser-ui-audit\managed-folder-transition-20260716-2021`.

### Build And ProgramData Deployment

- Revit 2019/2021/2023/2025/2027 Release build completed with zero compile errors. The existing `NETSDK1137` advisory remains for 2025/2027.
- Stage and ProgramData verification: `Installed OK` for all five targets; every Stage-versus-installed DLL comparison is `Match=True`.
- ProgramData backup: `_backups\programdata-before-managed-folder-homepage-transition-20260716-161023`.
- Deployment log: `artifacts\family-browser\programdata-install-managed-folder-transition-20260716-161032.log`.
- Installed DLL SHA256: 2019/2021/2023 `1B9A5EEFD6D24296EF3AEECBACD5FCBFC4F97C2302CDE8DE8C5DF0DDFC91ABE3`; 2025 `3B0B91845B73133586368735E769673BDF246D192D73E08253C41A6E556EFE41`; 2027 `F75E1CB0BEA343838E9292E4EF595E7040A57F494B4E1BC67B2F2946A6555AF6`.
- No installer or mail package was requested or generated for this change.

### Needs Revit Check

- On a PC with a persisted TEST override, publish a valid homepage management path, press Refresh, and confirm the TEST folder remains active while the homepage-return panel appears.
- Test `홈페이지 경로로 변경`, restart Revit, and confirm the homepage folder remains active instead of reverting to the TEST override.
- With disposable management data, test `기존 데이터 이관 후 변경`; verify the TEST source remains intact, copied homepage data is readable after restart, and a deliberate destination conflict blocks the migration without overwriting central data.
- Real network data was not migrated automatically during development. Migration behavior was verified only against isolated temporary fixtures.

## 2026-07-16 TEST Management Folder Transition Installer And Mail Package 16:17

- Previous latest-artifact pointer backup: `_backups\installer-mail-managed-folder-transition-20260716-161532`.
- Ran `Build-FamilyBrowserInstaller.ps1 -Version 1.0 -Label latest-20260716-161532 -MailPackageMinimumMB 15.5 -PreserveExistingArtifacts`; all earlier installer and mail-package artifacts were preserved.
- Revit 2019/2021/2023/2025/2027 Release builds completed with zero compile errors. The existing `NETSDK1137` advisory remains for 2025/2027.
- Stage and installed-payload verification passed for all five targets. Stage-versus-ProgramData DLL SHA256 comparison is `Match=True` for every target.
- Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260716-161532_Setup.exe`; size `3,597,182` bytes; PE `MZ` header PASS; SHA256 `9556D9BA25EDDD302E3F782FAED5BFE9B1E7D4D2218E24DE6900E8F35B8BD6C1`.
- Mail package: `artifacts\family-browser\mail-packages\20260716_06.zip`; size `16,777,451` bytes (`16.00 MB`); SHA256 `029E0FB1F38772EFD3488290E9B112DAA8DAE96B25814A45AD298A256500EF70`.
- ZIP entries are `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`. The embedded `Setup.exe` SHA256 exactly matches the standalone installer.
- Updated `latest-build.json`, `latest-mail-package.json`, and the ZIP checksum file to the validated artifacts. No Desktop copy was created.
- Build log: `artifacts\family-browser\build-installer-managed-folder-transition-20260716-161532.log`.

## 2026-07-18 Standard RVT Change Notification And Audit Review

### Finding

- Current implementation is only partial and is not yet a deployment-grade standard RVT audit trail.
- The header `Update Available` count is the current-project comparison result, and the `Tracking` pill is the current-project scan/stamp state. Neither is a live notification that the registered standard RVT file changed.
- Home `Recent Activity` shows saved snapshot/comparison/preflight/request/policy dates. It does not show the registered RVT's current file stamp or a detailed standard change history.
- Standards includes `Fast Stamp Check`. Only this explicit action calls `BuildStandardRvtFileStampStatus(..., probeFile: true)` and shows a structured prompt when the source file timestamp differs, offering a fast snapshot rebuild.
- Normal Home/Standards rendering calls `BuildStandardRvtFileStampStatus(registration)` with probing disabled. Therefore the existing `RVT updated` warning branch cannot become active during ordinary rendering. `IsStandardSnapshotStale(...)` exists but has no call site.

### Existing Tracking Coverage

- When the exact registered standard RVT is open in Revit with this add-in active, `DocumentChanged` collects added/modified loadable families, family types, references, and supported system types. `DocumentSaving` also compares the document with the last snapshot to find added/deleted families, family types, and system types.
- Candidate JSON stores Windows domain/user, UTC time, document path, trade/source, category, family/type, `Added`/`Modified`/`Deleted`, reason, and details under managed `StandardChangeCandidates`.
- The visible `Candidate Log` summary exposes only candidate count plus family/system names. It does not show user, time, change kind, or reason.
- Browser-driven Family load and System Type apply operations are separately queued and written to managed `OperationLogs` only after successful project Save/Save As/Synchronize. These operation logs are not currently exposed as a browser history view.

### Gaps Before Deployment

- No automatic lightweight source-file probe runs on browser open or top Refresh, so stale standard snapshots can continue to be displayed until an administrator manually runs `Fast Stamp Check` or a scan.
- File-change classification uses timestamp delta only. File length is calculated but not included in `ChangedSinceScan`, and no content hash is checked.
- A mapped drive, UNC path, and IP share path are normalized as strings but not resolved to one canonical file identity, so the same RVT opened through an alias can miss candidate tracking.
- Candidate entries are written during `DocumentSaving`, before successful save status is known, so a failed save can leave a false candidate. Repeated equivalent candidates replace the older entry, so this is a latest-candidate cache rather than an immutable history.
- Changes made without this add-in active, file replacement outside Revit, or a path-alias mismatch can be detected only as a manual file-stamp difference; the responsible user and exact changed items cannot be reconstructed reliably.
- Current automated QA verifies routes/event hooks and project-operation commit state, but it does not execute a real registered-standard-RVT edit/save/fail-save/external-replacement matrix.

### Recommended Next Work

- Add a cached, asynchronous file probe on browser open and top Refresh; compare timestamp, length, and a revision/hash manifest without opening the RVT.
- Show a persistent per-trade warning in Home, header status, Family/System lists, and Standards. Mark the snapshot stale and prevent stale standard data from being treated as current until a scan is accepted.
- Add a detailed change-history table with user, Revit user, machine, time, save/sync state, change kind, category, family/type, and before/after fingerprint summary.
- Commit standard change events only after successful Save/Save As/Synchronize, preserve immutable history, and canonicalize local/central/mapped-drive/UNC/IP identities.
- Add isolated file-stamp tests plus Revit runtime fixtures for family add, type add/delete/modify, system-type modify, failed save, sync, external replacement, and multi-PC shared-log visibility.

### Status

- Superseded by the implementation and deployment record below. Automatic source revision blocking and successful-commit history are fixed; the real Autodesk save/sync/external-replacement matrix remains `Needs Revit Check` in Next Work Queue item 32.

## 2026-07-18 Standard RVT Revision Blocking And Immutable History Implementation

### Fixed Behavior

- Added shared `FamilyBrowserStandardRevisionService` and per-source manifests under managed `StandardRevisionManifests`. Accepted registration/precise-scan operations record timestamp, length, sampled content hash, canonical path, and Windows file identity through atomic replacement.
- Browser startup preload, explicit Refresh, and a throttled 60-second background task probe registered sources without opening the RVT or blocking the UI thread. The active source state is keyed by source ID and standard slot; policy-store timing gaps use the in-memory registration only for the resolved active slot.
- Changed, missing, inaccessible, baseline-missing, or probe-error states now appear in the header and a per-trade Home board. Standards shows current/accepted stamps, identity information, reason, and confirmed history. Korean/English state text is refreshed on language change.
- A blocked source immediately clears stale Family/System rows and marks scan required. Current Model Check, Family load, and System Type apply validate the same revision service again immediately before mutation, so cached data cannot modify a model after the standard source changes.
- Source identity resolves mapped drives to UNC where possible and compares volume/file IDs/final paths. The sampled hash reads the first/middle/last 1 MB, so same-size and restored-timestamp content changes remain detectable without hashing a full large RVT on every probe.
- Standard candidate records now carry Windows/Revit user, machine, canonical source identity, category/family/type, change kind, and before/after fingerprint summaries. Save prepares pending records only; successful Saved/SavedAs/Synchronize commits immutable per-entry history. Failed/cancelled saves do not commit. Save As is recorded as source history only when the final document identity is the registered standard source.
- Immutable history lives under `StandardChangeCandidates\History\<source>\<date>` with a bounded read cache. Equivalent later saves are preserved instead of replacing an earlier event. Existing browser-driven operation logs also feed successful commit history.

### Automatic Coverage

- Added contract scenarios `admin-home-standard-rvt-changed` and `modeler-family-standard-rvt-unavailable` in Korean and English, plus stale-row blocking, Home board, status-pill, result-dialog, and language-purity assertions.
- Added a real 4 MB revision fixture that modifies only the middle byte while restoring timestamp and length; the sampled hash must change. Hardlink aliases must have the same file identity, while a same-content copy must have a different identity.
- Static checks lock baseline recording, startup/background probes, source-revision mutation guards, immutable successful-commit history, identity/hash code, contract scenarios, and primitive fixture assertions.
- Targeted all-runtime report: `artifacts\family-browser-standard-revision-audit\20260718-134147-all`; 2019/2023/2025/2027 all passed, 2021 runtime expected `SKIP runtime-not-installed`.
- Full quality gate: `artifacts\family-browser-ui-audit\20260718-134334-standard-revision-quality-gate`; elapsed `851.9s`, zero failures. Static/action contract counts are `80` generated actions, `257` exact routes, `63` prefix routes, and `11` browser functions for each unique host source set. Five-target Release build/Stage, Stage verification, nested-family propagation, authoritative System Type apply, 2,000-row performance/cache, and all installed-runtime Korean/English DOM/click/detail/layout scenarios passed.

### ProgramData And Distribution

- Revit process count before deployment: `0`. The first normal ProgramData copy was rejected by Windows ACLs; the corrected elevated install completed for Revit `2019`, `2021`, `2023`, `2025`, and `2027`.
- Installed manifest/payload verification passed for all five targets. Final Stage-versus-ProgramData DLL hashes match: 2019/2021/2023 prefix `4D950354FFBDFDDC`, 2025 prefix `94DF38FABFAC9383`, 2027 prefix `0BE7647D57761761`.
- Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_latest-20260718-140036-standard-rvt-tracking_Setup.exe`; `3.45 MB`; PE `MZ` PASS; SHA256 `EC23A9BA5A36E46AE64AA31779F510364F70ACFF3F8CF1F946AC6397F37A5A99`.
- Mail package: `artifacts\family-browser\mail-packages\20260718_01.zip`; `16.00 MB`; SHA256 `C6BC781F1BD61DA2DFE4F823A0EBFA5BB25DD0B12CA2F188E723AB3F647F519E`. Embedded `Setup.exe` exactly matches the standalone installer.
- `latest-build.json`, `latest-mail-package.json`, and `20260718_01.zip.sha256.txt` now point to and validate these final artifacts.

### Truth Boundary And Needs Revit Check

- A source file changed outside an active Family Browser session is still detected and blocked, including same-size/same-timestamp replacement, but its author and exact changed elements cannot be reconstructed. The UI must report source changed/unknown history, not infer a person or item list.
- Exact user/item history is authoritative only when this add-in observed the registered source and Revit reported successful Save, Save As, or Synchronize. Execute Next Work Queue item 32 with a disposable standard RVT before treating the Autodesk event matrix as production-proven.

## 2026-07-18 Project File Tracking And Lightweight Catalog Delta Review

### Current Coverage

- Browser-driven Family load and System Type apply operations are queued and committed to managed `OperationLogs` only after Revit reports a successful Save, Save As, or Synchronize. The commit path rechecks that the loaded/applied item is still present. This gives an authoritative record for mutations performed while this add-in is active.
- Saved Current Model Check data is stored per project identity. Workshared local files resolve to the central-model path when available. A changed project file timestamp or length invalidates the saved detailed scan instead of silently reusing it.
- The existing fast project scan enumerates `Family` names/categories and selected standard-supported `ElementType` names. It classifies items outside the approved Excel/JSON standard list as unregistered.

### Gap

- There is no accepted catalog baseline for the same project, so an externally loaded approved family does not produce a project-change warning.
- The current light family record does not include `FamilySymbol` type names. Type additions, deletions, and renames inside an existing family are therefore not detected by the fast scan.
- Deleted families are not reported as a delta from the project's prior state, and the light System Type scan is limited to classes present in registered standard snapshots.
- Startup intentionally uses saved project rows and skips current-document enumeration while `_deferExpensiveStartupLookups` is active. Therefore the current fast scan is not a guaranteed first-open detector.
- When the add-in was absent during a change, the later observer can identify added/removed catalog names but cannot truthfully reconstruct the responsible user, exact time, or edit mechanism.

### Recommended Project Catalog Baseline

- Add a separate per-project `project-family-catalog-v1.json`, keyed by canonical central-model identity with local/document aliases. Store sorted/deduplicated keys for category + loadable family, category + family + type, and class + category + supported System Type.
- Keep two states: an automatically updated `last observed` catalog for detecting what changed since the previous browser observation, and an `accepted baseline` for unresolved governance warnings. Persist immutable timestamped catalog revisions so updating `last observed` does not erase the audit trail.
- Capture only names and stable identifiers. Do not call `EditFamily`, open an RVT/RFA, generate fingerprints, inspect parameters/CSV, render thumbnails, or capture 3D images.
- Compare the current catalog on first browser open, explicit Refresh, and successful Save/Sync. Show a persistent summary such as `프로젝트 구성 변경 감지: 패밀리 +3 / -1, 타입 +7 / -2` plus a review table of added/removed names.
- Match current deltas against committed browser `OperationLogs`. Matching entries are known browser changes; unmatched entries are `외부/미추적 변경`. Do not infer an author for unmatched entries.
- Do not overwrite the accepted baseline as soon as a difference is detected. Update it only after a successful Current Model Check or an explicit administrator action such as `변경 확인 및 기준 갱신`; otherwise the warning would disappear before review.
- The first-ever catalog capture establishes `기준선 없음` and requires Current Model Check or administrator acceptance, rather than reporting every existing family as newly added.
- A rename is fundamentally an added + removed name unless a trustworthy persistent Revit element identity proves continuity. The primary UI should therefore report additions/removals and may show a rename candidate only as a non-authoritative hint.

### Performance Boundary

- The proposal is materially cheaper than Current Model Check because it consists of `FilteredElementCollector` enumeration plus name/category/type extraction and a sorted JSON/hash comparison.
- Revit element enumeration must execute on the Revit API thread. Sorting, hashing, persistence, and diff formatting can run after capture without holding the API context.
- Exact duration depends on project size and workstation. Treat `fast` as a measured contract, not an unconditional promise: log elapsed milliseconds and element counts, target sub-second capture for ordinary projects, and add a large-project runtime benchmark before enabling automatic first-open warnings.

### Status

- Superseded by the implementation and deployment record below. The source and installed payload now include the lightweight catalog tracker; the real Revit save/sync/external-change matrix and large-project timing remain `Needs Revit Check` in Next Work Queue item 33.

## 2026-07-18 Project Catalog Baseline Implementation And Deployment

### Implementation

- Backup: `_backups\project-catalog-tracking-20260718-154307`.
- Added shared `FamilyBrowserProjectCatalogService` and dashboard integration. Each project is keyed through the existing canonical project identity path, which resolves a workshared local model to its central-model identity when Revit supplies one.
- The lightweight capture enumerates loadable `Family` records, every `FamilySymbol` type, and the supported System Type classes. It does not call `EditFamily`, open RVT/RFA files, calculate fingerprints, inspect parameters/lookup CSV, render thumbnails, capture 3D images, or mutate the model.
- Managed `ProjectCatalogs` storage keeps atomic accepted and observed manifests plus immutable revisions. The first capture reports `기준선 없음`; later automatic observations update only last-observed state and cannot clear an unresolved accepted-baseline difference.
- Browser startup/activation, top Refresh, and successful Save, Save As, or Synchronize invoke observation. A successful full Current Model Check accepts the already-captured project catalog. Administrators additionally have an explicit name-only `변경 확인 및 기준 갱신` action with a warning that it is not fingerprint equality.
- Added a persistent header state and a Home project-composition board showing family/type/system additions and removals, item-level differences, capture timing/counts, and known Browser versus external/untracked classification. Added/loaded items matching committed managed `OperationLogs` are known Browser changes; unmatched additions and all unmatched removals remain external/untracked. The implementation does not invent an author or precise change time for edits made without the add-in.
- A rename remains represented as one removed name plus one added name unless Revit provides a trustworthy persistent identity. This avoids presenting a guessed rename as fact.
- Added exact actions `project-catalog-observe-auto`, `project-catalog-check`, and `project-catalog-accept`, plus audit scenario inputs and contract scenarios for administrator baseline-missing and normal-user external/untracked-change states.
- Extended `Invoke-FamilyBrowserUiAuditHarness.ps1` with `-ScenarioNamePattern` so a changed contract surface can be rerun quickly without waiting for every unrelated dashboard scenario.

### Bug Found During Automation

- The first full harness run exposed an actual `DateTimeStyles` bug in the new audit-state parser: `RoundtripKind` was combined with `AssumeUniversal`/`AdjustToUniversal`, which .NET rejects. This made every Home scenario touching project-catalog state fail even though compilation passed.
- Replaced the invalid combination with `AssumeUniversal | AdjustToUniversal` in the service and dashboard parser, rebuilt every target, and reran the focused Korean/English IE scenarios successfully.

### Automated Verification

- Revit 2019/2021/2023/2025/2027 Release builds completed with zero compile errors. Existing generated-source/platform analyzer warnings remain unrelated to this change.
- Static/action/contract: PASS for all three host source sets with `80` generated actions, `260` exact routes, `63` prefix routes, and `11` Browser functions each.
- Targeted IE `WebBrowser` report: `artifacts\family-browser-ui-audit\20260718-162051`. Administrator baseline-missing and normal-user external/untracked-change scenarios passed in Korean and English for 2019/2023/2025/2027. Revit 2021 runtime is not installed and is recorded as `SKIP runtime-not-installed`; its shared 2019-2023 binary, contract, build, and Stage passed.
- Integrated report: `artifacts\family-browser-ui-audit\20260718-162230-quality-gate`. Static/contract, nested-family difference propagation, authoritative System Type apply, five-target Stage generation/verification, and the 2,000-row performance/cache gate all passed.
- The earlier all-scenario gate was intentionally stopped after its ten-minute command limit while progressing normally. The focused harness option was added so this did not become a false product failure; the failing DateTime parser found in its partial artifacts was fixed before the final focused and integrated passes.

### ProgramData Deployment

- Revit process count before deployment: `0`.
- ProgramData backup: `_backups\programdata-before-project-catalog-20260718-162416`.
- The first ordinary copy was denied by the ProgramData ACL. The first hidden elevated attempt did not start and produced no log or partial deployment. A visible Windows UAC elevation then completed with exit code `0`.
- Installed manifest/payload verification is `OK` for Revit 2019, 2021, 2023, 2025, and 2027.
- Stage-versus-ProgramData DLL SHA256 comparison is `Match=True` for all targets: 2019/2021/2023 `354E0B8501319605EC435D0AD9308AC789A68B3AC8F662D9F4077C05DF9ADD0B`; 2025 `A1D63259FA0CF6E1B6EE9EBB92ACA2DB4293836D425824AB0A98BF9D4FA9345E`; 2027 `5548F22BD810D3CCD152E287ED20C1B637E6151BECA1039B590DE3107FCD9B6C`.
- No installer or mail package was requested or generated for this change.

### Needs Revit Check

- In a disposable workshared project, accept the first baseline, load a Family through the Browser and Save/Sync, then verify the delta is classified as a known Browser change until accepted.
- With the add-in absent or disabled, add/delete/rename one Family and one Family Type, reopen with the add-in, and verify persistent added/removed external/untracked warnings. Confirm that merely observing the project again does not clear the accepted-baseline warning.
- Exercise failed/cancelled Save and successful Save As/Sync, local and central aliases, and two PCs opening the same central project identity.
- Measure `CaptureElapsedMilliseconds` on a genuinely large production-like project. The capture is intentionally name-only and expected to be much cheaper than Current Model Check, but its real Revit API-thread duration is not claimed as sub-second until this fixture is run.

## 2026-07-18 End-to-End Workflow, Tracking, Storage, And Long-Path Hardening

### Scope And Backup

- Source checkpoint: `_backups\workflow-lifecycle-hardening-20260718-180946`.
- ProgramData checkpoint before the attempted final deployment: `_backups\programdata-before-workflow-deep-audit-20260718-193412` (`17` files).
- Audited the full browser lifecycle contract: startup and managed-folder selection, online/offline cache, Home state, Family/System browse and detail, request workflow, administrator live refresh, standard registration/list connection, Current Model Check, Family load/System apply, Save/Save As/Sync commit, project name-only baseline, nested Family propagation, diagnostics, language/layout, and migration between TEST and homepage-managed roots.

### Problems Found And Fixed

- Fixed invalid `DateTimeStyles` combinations in standard-revision, dashboard, and tracking parsers. The previous `RoundtripKind` plus assume/adjust flags could throw at runtime while compilation still passed.
- Added durable tracking write-ahead persistence under `%LOCALAPPDATA%\KKY\FamilyBrowser\PendingTracking`. Managed operation/history writes now use stable IDs, immutable history files, reconnect flush, and idempotent replay. Home and the header expose pending offline tracking instead of silently hiding it.
- Tightened Family load tracking to the exact Family Type. A same-family/different-type match can no longer incorrectly verify and commit the requested operation.
- Fixed project-catalog ExternalEvent retry throttling. The throttle timestamp is now advanced only after dispatch succeeds; a failed dispatch no longer suppresses the next legitimate retry.
- Added `ProjectCatalogs` and `StandardRevisionManifests` to managed-folder bootstrap and migration in all three host families. TEST-root to homepage-root migration now includes these governance records and rebases migrated JSON paths.
- Hardened request-store writes in all three hosts. They no longer delete the committed JSON before promoting the new copy, and a failed promote restores the prior committed file.
- Fixed the same delete-before-promote data-loss window in all three standard-list writers. An approved Excel/JSON list now survives a network interruption during replacement.
- Added shared `FamilyBrowserAtomicFileService` with same-folder temporary files, native `File.Replace`, backup/restore fallback, and recoverable promotion. Data loader caches, measurement preference, nested-only catalog, project catalog, standard revision, tracking, managed-folder override, request store, and standard-list writes now use the hardened path.
- Found a genuine deep-path failure in the Revit 2023 performance gate. A destination path of `215` characters became `266` characters after the old `.kky-migration-<32-char-guid>.tmp` suffix and exceeded classic .NET Framework path limits. Temporary/backup/probe names are now short `.kky-t-<8>.tmp`, `.kky-b-<8>.bak`, and `.kky-w-<8>.tmp` sibling names; migration also ignores these auxiliary files.
- Expanded the workflow contract from `19` to `26` workflows and added startup-cache-offline, live administrator refresh, Family/System detail integrity, nested-family propagation, large-library performance, diagnostics, and tracking persistence coverage.
- Removed a UI-audit false pass by forwarding project-catalog baseline/change/untracked scenario inputs into the actual dashboard DOM and validating the resulting state in the IE harness.

### Automated Verification

- Static/action/contract checks: PASS for all three host source sets. Each source set reports `80` generated actions, `260` exact routes, `63` prefix routes, and `11` browser functions.
- Workflow report: `artifacts\family-browser-workflow-audit\workflow-deep-audit-long-path-fix`; PASS, `26` workflows and `271/271` checks.
- Core quality gate: `artifacts\family-browser-ui-audit\20260718-workflow-deep-audit-core-pass\quality-gate-summary.md`; PASS for static/contract, workflow, managed-data audit, five-target build/Stage verification, and performance.
- IE `WebBrowser` scenario runs: `harness-rvt2019`, `harness-rvt2023`, `harness-rvt2025`, and `harness-rvt2027` under `artifacts\family-browser-ui-audit\20260718-workflow-deep-audit-final2`; each passed `52/52`, total `208/208`. Revit 2021 runtime is not installed, so it remains `SKIP runtime-not-installed`; its shared 2019-2023 binary, contract, build, and Stage passed.
- The 2,000-row Family/System cache and interaction performance gate passed on installed runtimes. The repaired Revit 2023 long-path case also passed at `artifacts\family-browser-ui-audit\20260718-performance-rvt2023-long-path-fixed`.
- Revit 2019/2021/2023/2025/2027 Release build and Stage verification passed with zero compile errors. Final Stage DLL SHA256 values are: 2019/2021/2023 `A6A2948F1E81B7079E360F50E304943E9E38DAF1AE2ADEA178F4E318643D254D`; 2025 `44F3CBEE923C79D0EA4A0C4AD9D81C75F5B4AADB08205F96FBE1DDAB9F5C058C`; 2027 `589A9DC87AC9CD4A053DB64133427DBA8F2E7F58A0EC681DF1C41FB894C9CBD6`.
- Homepage bootstrap responded HTTP `200` with policy version `2026.05.19-kcim-test`, but both configured candidates (`I:\30. 협력사 전용폴더\00. BIM_KCIM\02. 패밀리\TEST` and `D:\TEST`) are unavailable on this PC. Therefore real managed V2 manifests/row caches could not be inspected here; the report correctly records environment `UNAVAILABLE` rather than product PASS.

### ProgramData Deployment Status

- Revit process count was `0`. The ordinary ProgramData copy was denied by ACL, and both autonomous `RunAs` attempts ended with Windows reporting that the user cancelled the operation. No UAC/security dialog was automated.
- The failed ordinary install stopped on the first DLL, so no mixed partial payload was left. The installed payload remains the previous verified build.
- Installed DLL SHA256 values remain 2019/2021/2023 `354E0B8501319605EC435D0AD9308AC789A68B3AC8F662D9F4077C05DF9ADD0B`; 2025 `A1D63259FA0CF6E1B6EE9EBB92ACA2DB4293836D425824AB0A98BF9D4FA9345E`; 2027 `5548F22BD810D3CCD152E287ED20C1B637E6151BECA1039B590DE3107FCD9B6C`. These do not match the final Stage hashes above.
- Status: source and Stage are verified; final ProgramData deployment is pending one accepted Windows elevation. Do not claim this final revision is active in Revit until the post-install hash comparison is `Match=True` for all five targets.

### Truth Boundary

- Automated tests prove routing, rendered conditions, storage recovery, cache/migration behavior, deterministic tracking persistence, and Revit-version compilation. They do not prove Autodesk runtime callbacks or model mutation against a real workshared fixture.
- Complete Next Work Queue items `32`, `33`, and `34` before production-signing the Save/Sync/external-edit/two-PC lifecycle. Failed/cancelled operations must create no committed history; unmatched external changes must remain author-unknown rather than inventing attribution.
- Superseded by the request-concurrency implementation below. Atomic files plus request-scoped locks and optimistic revision/token checks now reject stale status and delete actions; real two-PC SMB behavior remains Next Work Queue item `37`.

## 2026-07-18 Request Optimistic Concurrency And Multi-Writer Protection

### Scope And Backup

- Source checkpoint: `_backups\request-optimistic-concurrency-20260718-201206`.
- ProgramData checkpoint: `_backups\programdata-before-request-concurrency-20260718-202630` (`17` files).
- Scope was intentionally limited to request create/update/status/delete persistence, dashboard request actions, managed-folder migration rules, workflow automation, and deployment. Revit model mutation logic was not changed.

### Problem Found And Fixed

- The existing atomic replace prevented a damaged or half-written request JSON, but two administrators could still render the same revision and overwrite one another semantically. Status and delete URLs carried only the request ID, so the server could not tell whether the browser screen was stale.
- Added shared `FamilyBrowserRequestConcurrencyService` with a deterministic request-scoped same-folder lock file. A mutation obtains the lock with `FileShare.None`, re-reads the authoritative request, validates the browser's expected revision and token, and only then writes or deletes.
- Added `Revision` and `RevisionToken` to request records in all three host families. Every successful update advances the revision and generates a new token. Existing records without these fields use a SHA-256 content token, so a change written by an older client before the new action is still detected.
- Request status and delete links now carry request ID, expected revision, and expected token. Incomplete legacy routes are rejected without mutation. A stale or lock-busy action shows a structured conflict result and reloads the latest request list instead of silently overwriting it.
- Request lock files use `.kky-r-<hash>.lck` names and are excluded from managed-folder migration. The lock file remains as an empty coordination object after release, avoiding a delete/recreate race between waiting writers.

### Automated Verification

- Workflow contract expanded from `26` to `27` workflows. Report: `artifacts\family-browser-workflow-audit\request-concurrency-20260718\workflow-audit-summary.md`; PASS, `285/285` checks.
- The executable fixture proves same-request lock exclusion and clean reacquisition, stale-revision rejection, and legacy-file content-token rejection. Static QA verifies lock/revision/token integration across Revit 2019-2023, 2025, and 2027 sources.
- Focused IE `WebBrowser` request/permission scenarios passed for installed Revit 2019, 2023, 2025, and 2027 runtimes at `artifacts\family-browser-ui-audit\20260718-request-concurrency-focused`; all visible request actions emitted valid host routes with zero failures. Revit 2021 remains `SKIP runtime-not-installed`, while its shared binary build and Stage passed.
- Integrated quality gate: `artifacts\family-browser-ui-audit\20260718-request-concurrency-quality-gate\quality-gate-summary.md`; PASS for static/action contract, workflow, nested-family propagation, authoritative System Type apply, five-target Stage verification, and the 2,000-row performance/cache gate. Managed live data was correctly recorded as environment `UNAVAILABLE` because this PC cannot reach the configured homepage candidates.
- Revit 2019/2021/2023/2025/2027 Release build and Stage verification passed with zero compile errors. Final Stage DLL SHA256 values are: 2019/2021/2023 `ABD94123F82F31B16DD36AEA8E99B84E15EE9FFBD771E7C637103A2B3AF901C6`; 2025 `A248D62791B98C877354A3A33FC83F5AB6BC71015B6AC57702131503EB004F39`; 2027 `899F235D3AD1F70475C5128AC95BF72800950FE61294EE07B48A833474CF9189`.

### ProgramData Deployment

- Revit process count before deployment was `0`. The ordinary copy was denied by ProgramData ACL, then one normal Windows elevation completed successfully with exit code `0`; no UAC/security dialog was automated.
- Installed DLL and `.addin` manifest hashes match Stage for all five targets. ProgramData DLL SHA256 values exactly match the Stage values above.
- Install log: `artifacts\family-browser\programdata-request-concurrency-install.log`.

### Truth Boundary And Remaining Hardening

- The lock is authoritative when every PC writes the same real shared SMB folder, including mapped-drive and UNC aliases that reach the same server object. It is not a distributed lock for cloud-synced or eventually consistent local replicas.
- A participating old client does not honor the new lock. Its completed earlier write is detected by the legacy content token, but a truly simultaneous old-client write cannot be guaranteed safe. Update every workstation before relying on concurrent request edits.
- Automated fixtures prove persistence semantics without Revit, but Next Work Queue item `37` must exercise the conflict UI and shared-file behavior from two real PCs.
- Attachment copy idempotency and immutable deleted-request history were completed in the subsequent implementation below.

## 2026-07-18 Request Attachment Transaction And Immutable Deletion Audit

### Scope And Backup

- Source checkpoint: `_backups\request-attachment-audit-20260718-203407`.
- ProgramData checkpoint: `_backups\programdata-before-request-attachment-audit-20260718-204236` (`17` files).
- This change affects request attachment persistence and request deletion audit only. Family/System scanning, comparison, load/apply, and Revit model mutation paths were not changed.

### Problems Found And Fixed

- Request attachments were copied with incrementing filenames before the authoritative request JSON was committed. If JSON, mail draft, manifest, or request-store fallback failed, a retry could copy the same source again and could retain attachment metadata that pointed to the previously failed store.
- Added shared `FamilyBrowserRequestFileTransactionService`. It reads each selected source once while simultaneously hashing and copying to a short same-folder temporary file, then promotes it to a deterministic `safe-name + SHA256 prefix` path. Same-content retries reuse the existing validated file; changed content receives a different path.
- `FamilyBrowserRequestStore.SaveLocked` now snapshots original attachment metadata, tracks files created by the current attempt, and treats the atomic request JSON promotion as the authoritative commit point. A pre-commit failure removes newly created files, restores both attachment lists and the original folder, and removes an abandoned request-specific folder for a new request. This also keeps fallback to a different request store clean.
- Mail draft, request-store manifest, and attachment manifest are auxiliary after the authoritative JSON commit. Their failure no longer makes the UI retry an already committed request and duplicate attachments. The Browser shows a structured saved-with-warning result, records the technical exception in diagnostics, and exposes details/log path only to administrators.
- Attachment metadata now persists `ContentSha256` alongside size and stored path.
- Request deletion previously removed the active JSON before attachments and erased the only embedded history. It now records an immutable full `DeletePrepared` event under `RequestAudit\Deleted\yyyyMMdd` while holding the request lock, then removes attachment data and mail draft before deleting active request JSON. Success writes `DeleteCompleted`; a cleanup exception writes `DeleteCleanupFailed` and leaves the active request JSON until cleanup can be retried.
- Deletion audit events preserve request ID, revision/token, source-file token, actor, time, admin-delete state, attachment count, and the complete request snapshot including history and attachment metadata. Immutable creation uses same-folder temporary writing plus create-new final promotion and cannot overwrite an earlier event.

### Automated Verification

- Static/action/contract checks passed for all three host source sets with unchanged route counts: `80` generated actions, `260` exact routes, `63` prefix routes, and `11` browser functions.
- Workflow contract expanded from `27` to `28` workflows. Final report: `artifacts\family-browser-workflow-audit\request-attachment-audit-final-20260718\workflow-audit-summary.md`; PASS, `300/300` checks.
- The executable fixture proves identical-content retry reuse, changed-content path separation, pre-commit created-file rollback, and immutable audit overwrite rejection while preserving the prepared event.
- Revit 2019/2021/2023/2025/2027 Release build and Stage verification passed with zero compile errors. Existing generated-source/platform warnings are unrelated.
- Integrated quality gate: `artifacts\family-browser-ui-audit\20260718-request-attachment-audit-final-quality-gate\quality-gate-summary.md`; PASS for static/action contract, the 28-workflow lifecycle audit, nested-family propagation, authoritative System Type apply, five-target Stage verification, and the 2,000-row performance/cache gate. Live managed data remains correctly reported as environment `UNAVAILABLE` on this PC.

### ProgramData Deployment

- Revit process count before deployment was `0`. Elevated install completed with exit code `0`; no UAC/security dialog was automated.
- Installed DLL and `.addin` hashes match Stage for Revit 2019, 2021, 2023, 2025, and 2027.
- Final DLL SHA256 values: 2019/2021/2023 `C874BCAF1CC59782796DD9A397ABE15F07BD2A23204AA378FD2BCB429DAE2F10`; 2025 `060D4DF2D4ED79998A795C7556F8EA4997F0D7CB42392E352BD6FF434FA53ABF`; 2027 `4E0B6E1B4C728CE7CFC556AC75E7CA58C5A887F0A87C630F0E7E677C43974338`.
- Install log: `artifacts\family-browser\programdata-request-attachment-audit-final-install.log`.

### Truth Boundary

- The file fixture proves deterministic local/SMB-compatible semantics, but it does not simulate a real server disconnect at each byte boundary. Execute Next Work Queue item `39` before calling large-attachment network interruption recovery production-proven.
- `DeletePrepared` is mandatory and deletion aborts if that history cannot be preserved. `DeleteCompleted` is supplementary; if its write fails after active deletion, the prepared full snapshot still prevents history loss.
- The Browser does not yet expose a dedicated deleted-request-history screen. The immutable records are retained for governance and diagnostics without adding another visible panel to the current Requests UI.

## 2026-07-18 Pre-Deployment Final Readiness Review

### Scope And Checkpoints

- Source/distribution checkpoint: `_backups\predeploy-distribution-gate-20260718-214700`.
- ProgramData checkpoint: `_backups\programdata-before-predeploy-final-20260718-215300` (`17` files, `19,809,094` bytes).
- Revit process count during final package and installed-payload verification: `0`.
- This pass reviewed source, Stage, ProgramData, installer, mail ZIP, metadata/checksum files, abandoned packaging work folders, static routes, workflow contracts, IE `WebBrowser` click/layout/language scenarios, and large-list performance.

### Problems Found And Fixed

- Found a real stale-distribution risk: the then-latest installer/mail package predated the final request-concurrency and attachment-transaction DLLs even though Stage and ProgramData were current. A fresh package was generated from the current Stage revision.
- Stage metadata previously listed only high-level build information. `stage-manifest.json` now records every payload relative path, byte length, and SHA-256 hash.
- Stage verification now validates the complete manifest. Installed verification compares every Stage file against ProgramData and rejects missing, changed, extra, or obsolete alternative add-in payloads.
- The direct ProgramData installer now checks for administrator elevation before mutation, blocks while Revit is running, copies and verifies a temporary payload first, removes the known old payload only after validation, and then promotes the verified payload.
- Both Inno Setup definitions now default to product version `1.0` and remove known obsolete payload/manifests for Revit 2019/2021/2023/2025/2027 before extraction.
- The installer builder now validates PE `MZ`, Stage revision metadata, standalone and embedded Setup hashes, mail-package contents/size, latest metadata, and cleans its own `mailpkg-*` work folder. Twelve abandoned generated work folders from older packaging runs were removed; final count is `0`.
- Added independent `Test-FamilyBrowserDistribution.ps1`. Its first run exposed a test-only PowerShell defect: successful `.ps1` calls were incorrectly judged through native-program `$LASTEXITCODE`, and failed checks prevented metadata from reaching later checks. It now uses exception-based script failure handling and an explicit shared metadata context. Static guards prevent that regression.
- The final path-form sanity check found that `Verify-FamilyBrowserRecovered.ps1` compared an absolute enumerated file path with a relative manifest path and could count `stage-manifest.json` as an eighteenth payload. Stage roots and manifest comparisons are now normalized to full paths; both relative-root and absolute-root verification pass.
- A non-elevated direct-install probe exposed the ProgramData ACL before any new application revision was required. Up-front elevation preflight was added. A later UAC attempt was cancelled by the user, but no application source changed in this distribution-only pass and the final full hash audit proves current ProgramData already exactly matches current Stage.

### Automated Verification

- Full quality gate: `artifacts\family-browser-ui-audit\20260718-predeploy-full\quality-gate-summary.md`; PASS in `915.1s`.
- Static/action contract: PASS for all three unique host source sets; each reports `80` generated actions, `260` exact routes, `63` prefix routes, and `11` browser functions.
- Workflow lifecycle: PASS, `28` workflows and `300/300` checks.
- IE `WebBrowser` HTML/click/layout/language harness: PASS with zero failures for Revit 2019/2023/2025/2027 Korean/English scenarios. Revit 2021 is `SKIP runtime-not-installed`; its shared 2019-2023 binary, Stage, contract, and package are verified. Harness warnings are only expected/handled empty-row selection alerts.
- 2,000-row performance gate: PASS. Maximum measured shell `13ms`, cold usable `719ms`, warm usable `423ms`, filter `17ms`, and cache operation `74ms`; rendered DOM remains windowed at `150` rows while all `1,000` rows per tab remain reachable.
- Final distribution audit: `artifacts\family-browser-distribution-audit\20260718-predeploy-final2-installed\distribution-audit-summary.md`; PASS for all `17` Stage payload files, installer freshness/revision, mail ZIP embedded Setup, zero abandoned work folders, and full ProgramData equality across five Revit targets.
- Relative-root regression audit: `artifacts\family-browser-distribution-audit\20260718-predeploy-final3-relative-root`; PASS with `-StageRoot .\artifacts\family-browser\stage` and installed-payload verification enabled.
- Tamper self-test: `artifacts\family-browser-distribution-audit\20260718-tamper-selftest`; PASS because a deliberately altered Stage payload hash was rejected as `Installer Stage revision differs`.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage verification completed with zero compile errors. Revit 2025/2027 retain only the existing `NETSDK1137` SDK-style advisory warning.

### Final Distribution

- Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_predeploy-final2-20260718-2201_Setup.exe`.
- Installer SHA-256: `927893BF96A98C7EE9DDFBBABFDB99DF4B01F6FB09C77A26DD88A36D1188D9A8`.
- Mail package: `artifacts\family-browser\mail-packages\20260718_03.zip` (`17,196,877` bytes, `16.4 MB`).
- Mail package SHA-256: `FC0F3C8DD204047AC94FDEBD31CDDFB30D85208DB7C8E9268A60E290F0D02308`.
- ZIP contains exactly `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`; embedded Setup exactly matches the standalone installer.
- `latest-build.json`, `latest-mail-package.json`, and their checksum files point to the artifacts above.

### Remaining Release Risks And Tomorrow's Revit Checks

- `Needs Release Decision`: the installer and add-in DLLs are Authenticode `NotSigned`. This does not change product behavior, but SmartScreen or enterprise endpoint policy may warn or block on another PC. Production code signing needs a trusted certificate and cannot be truthfully simulated with a self-signed test certificate.
- `Needs Managed-Share Check`: this PC received HTTP `200` from the homepage policy, but neither configured managed-folder candidate exists here. Live shared snapshots, real homepage-root migration, SMB aliases, offline/reconnect, and current production data content could not be audited; automation correctly reports environment `UNAVAILABLE`, not product PASS.
- `Needs Revit Check`: execute Next Work Queue items `32` and `33` with disposable standard/project RVTs. Verify successful versus cancelled Save/Save As/Sync, stale standard blocking, browser operation commit, external Family/type additions/deletions, accepted project baseline, and real large-project catalog timing.
- `Needs Two-PC/SMB Check`: execute items `34`, `37`, and `39`. Verify mapped-drive/UNC identity, request stale-edit/delete conflicts, large attachment interruption/retry, immutable delete audit, pending offline tracking flush, and TEST-root to homepage-root migration.
- `Needs Guard/Scan Check`: with Admin OFF, verify Family Load, Family/Type rename, nested-only placement, and file-specific guard rules in a real model. Run precise scan fixtures that produce OK-only and Delete/Cancel dialogs, including `Opening not cutting anything`, and confirm the automatic response plus on-demand XLSX diagnostics.
- `Needs Mutation Check`: apply one existing and one new Family/System Type, including layered types, Railing/Stair, routing preferences, materials, and Curtain Panel dependencies. Confirm results stay pending until successful Save/Sync, cancelled close does not commit, details/fingerprints match Revit, and no duplicate type/material is produced.

## 2026-07-19 Opt-In Project Element Creation/Modification/Deletion Ledger

### Scope And Checkpoints

- Source checkpoint: `_backups\element-change-tracking-20260719-005058`.
- ProgramData checkpoint: `_backups\programdata-before-element-change-tracking-20260719-014223`.
- This work adds an administrator-controlled project element activity ledger. It does not change Family/System comparison, scanning, Fingerprint, load/apply, or model mutation rules.
- The administrator setting is opt-in and defaults to OFF for existing and new policies.

### Implementation

- Added `프로젝트 요소 생성·수정·삭제 추적` / `Track project element creation, modification, and deletion` to the administrator Standards settings in all three host source sets. Route `kkyfb:project-element-change-tracking/{on|off}` is permission-checked, persisted immediately, and refreshes the live dashboard state.
- Added shared `FamilyBrowserElementChangeTrackingService` and connected Revit `DocumentOpened`, `DocumentChanged`, `DocumentSynchronizingWithCentral`, Save/Save As/Sync completion, document closing/closed, and shutdown events for Revit 2019/2021/2023/2025/2027.
- The service captures a project baseline, reduces added/modified/deleted IDs against that baseline, records transaction names, handles Undo/Redo activity, and suppresses incoming Synchronize-with-Central changes so another user's downloaded work is not falsely attributed to the current workstation.
- A workshared project is committed only after successful Synchronize with Central. A standalone project is committed only after successful Save or Save As. Failed/cancelled operations produce no immutable commit.
- Each committed item keeps the available element ID, category, family/type, level, workset, location, before/after state, transaction evidence, Revit user, Windows user, machine, commit type, and UTC time. The record also carries an explicit coverage note when observation began after the project had already been edited.
- Managed history is append-only under `ElementChangeHistory\<project hash>\<yyyyMMdd>\<commit>.json`. When the managed root is unavailable, the same commit is written first to `%LOCALAPPDATA%\KKY\FamilyBrowser\PendingTracking\ElementChanges` and replayed idempotently after reconnect.
- No XLSX/CSV is generated automatically. The implementation stores compact JSON evidence only.

### Problems Found And Fixed During Verification

- A cancelled `DocumentClosing` could previously discard the session before `DocumentClosed` proved that the document actually closed. Closing now only maps the document ID; session removal occurs on the completed close event.
- Revit 2027 rejected direct `ElementId.IntegerValue` access. The shared code now reads `Value` or legacy `IntegerValue` through a compatibility helper, allowing the same source to build across all five targets.
- The UI harness assumed exactly one administrator checkbox because only the optional System Type detail setting existed before this feature. It now requires and individually validates both settings, their routes, checked state, copy, and non-overlapping geometry.

### Automated Verification

- Final quality gate: `artifacts\family-browser-ui-audit\20260719-element-change-tracking-final-v2\quality-gate-summary.md`; PASS for static/action contract, workflow lifecycle, five-target Stage verification, 2,000-row performance/cache, and IE `WebBrowser` HTML/click/layout/language checks.
- Static/action contract: PASS for all three host source sets (`81` generated actions, `260` exact routes, `64` prefix routes, and `11` browser functions each).
- Workflow audit: `29` workflows and `352/352` checks passed. The new fixture proves offline element spooling, immutable per-project reconnect flush, idempotent replay, and complete retention of `20` parallel commits.
- HTML/IE harness: `208` scenarios passed, one expected Revit 2021 runtime skip, zero failures, `5,312` clickable controls inspected, `4,160` host-action candidates, and `1,072` browser-only clicks. The `33` warnings are expected handled empty-row alerts.
- Performance gate remained within target. Across installed runtime targets, 1,000-row Family/System scenarios rendered only 150 DOM rows; worst shell was `12ms`, cold usable `653ms`, warm usable `380ms`, and filter `14ms`.
- Revit 2019/2021/2023/2025/2027 Release build and Stage verification passed with zero compile errors. Revit 2021 remains build/Stage/install verification only because its runtime is not installed on this PC.

### ProgramData Deployment

- Revit process count before deployment was `0`. Elevated installation completed successfully and the installed payload passed complete Stage-to-ProgramData verification for Revit 2019, 2021, 2023, 2025, and 2027.
- Installed core DLL SHA256: 2019/2021/2023 `F8838A1E40935C9D10312DAAE654C480327234F73E3A93C3AF8764622CF3BAF3`; 2025 `1345D76E34C6E6797CACF2C32B489F94BAEE3EA59109A1848C48E1D5C263B31A`; 2027 `FB83D9B4A091B89169B7645EA2EDC583F1A4DFCE0CE76E2F1443157C80231A90`.
- Install log: `artifacts\family-browser\programdata-install-20260719.log`.

### Truth Boundary

- This ledger can provide exact per-user/per-transaction evidence only on workstations where this revision of the add-in is running. A workstation without the add-in cannot retroactively reveal which user created, modified, or deleted an element between observations.
- The existing lightweight project catalog can still detect later Family/type name drift from an uninstrumented workstation, but that is an unknown external change, not proof of the responsible user or exact intermediate synchronization event.
- A modified event stores useful element identity and before/after summary metadata; it is not a complete audit of every parameter value, geometry delta, or model operation. Elements unavailable after deletion may retain only their last observed metadata.
- The first enabled baseline enumerates project elements. Automated fixtures validate logic and performance guards, but real large-model baseline cost, Autodesk event ordering, cancelled/failed native operations, incoming central changes, and two-PC attribution remain `Needs Revit Check` in Next Work Queue item 40.

## 2026-07-19 Element Tracking Gap Audit

### Scope And Checkpoints

- Source checkpoint: `_backups\element-tracking-gap-audit-20260719-094812`.
- ProgramData checkpoint: `_backups\programdata-before-element-tracking-gap-audit-20260719-102125`.
- This pass audited whether the opt-in element ledger could misattribute incoming central work, convert an incomplete baseline into false history, lose deletion evidence, or become inaccessible after writing valid records.

### Problems Found And Fixed

- Incoming Reload Latest was not covered by the Synchronize-with-Central suppression path. Revit 2023/2025/2027 now attach compatible `DocumentReloadingLatest` and `DocumentReloadedLatest` delegates through a reflection bridge; Revit 2019 uses a conservative transaction-name fallback because its API does not expose those events. A completed external update rebases the session instead of recording the incoming changes as local work.
- Initial baseline capture could leave a partial dictionary after an exception and later classify the unseen remainder as newly created. Baseline capture now returns explicit success/failure, clears partial state, uses one collector pass, and refuses to create a tracking session when the baseline is not trustworthy.
- A deletion first observed without previous metadata was silently omitted. It is now retained with `PreviousStateUnavailable` evidence so the event is visible without inventing element details.
- Save, Save As, Sync, close, and related App handlers called several services inside one exception boundary. Each service is now isolated by `RunEventHandlerSafely`, so an unrelated policy/catalog exception cannot prevent the element ledger from committing or cleaning up.
- The immutable JSON history had no direct browser reader. Added administrator-only `최근 변경 이력 보기`, a structured HTML summary of the latest 200 commits and recent changes, local pending-queue visibility, and explicit XLSX export. Merely opening the viewer creates no workbook.

### Automated Verification

- Final quality gate: `artifacts\family-browser-ui-audit\20260719-element-tracking-gap-audit-final\quality-gate-summary.md`; all six stages passed: static/action contract, workflow lifecycle, five-target build and Stage, Stage verification, 2,000-row performance/cache, and full IE `WebBrowser` click/layout/language checks.
- Static/action contract passed for all three host source sets with `82` generated actions, `261` exact routes, `64` prefix routes, and `11` browser functions per host.
- Full IE harness result: `208 OK`, one expected Revit 2021 runtime skip, zero failures, `5,328` clickable controls inspected, `4,176` host actions, and `1,072` browser-only clicks. The `33` warnings are expected handled empty-row alerts.
- Performance remained within target with 1,000 Family plus 1,000 System rows and only 150 DOM rows per active list: worst shell `13ms`, cold usable `779ms`, warm usable `444ms`, filter `13ms`, and cache operation `85ms`.
- API metadata and delegate binding were inspected against installed Revit APIs: Reload Latest events exist and bind on 2023/2025/2027; Revit 2019 correctly takes the fallback path. Actual event ordering remains a Revit runtime check.

### Build And ProgramData Deployment

- Revit process count was `0`. Release build, Stage, and installed verification passed for Revit 2019, 2021, 2023, 2025, and 2027. Revit 2021 is build/Stage/ProgramData verification only because its runtime is not installed on this PC.
- ProgramData install log: `artifacts\family-browser\programdata-install-20260719-element-tracking-gap-audit.log`.
- Stage and installed core DLL SHA256 match: 2019/2021/2023 `0C2F3DBE81BAB42E1453D3834C7A369E736B9AF5B8A21BB85ED7CDA80623A057`; 2025 `ADCB4DBD1FEBD5AB0BBCB5EEBBE3D09A0C0E4C79781CFE53D41C29DDEFA52BEB`; 2027 `21075E5F162953DB04828AE72AE14BDFEF3E85868BFDCE29C21436A1240E84A7`.

### Remaining Truth Boundary

- Real two-PC SMB/central behavior, native Reload Latest event ordering, cancelled native operations, and large-model first-baseline cost remain `Needs Revit Check` in queue item 40.
- Work made on a PC without the add-in can be detected later only as an external coverage gap. It cannot be assigned retroactively to a precise user, transaction, or synchronization point.
- Client clocks do not establish a guaranteed global order, simultaneous edits to the same element are not merged into a forensic timeline, and append-only retention/archive policy is not yet fixed. These are `Needs Design` in queue item 41.
- The ledger records element identity and useful before/after summaries. It is not a complete parameter-by-parameter or geometry-delta audit of all Revit modeling activity.

## 2026-07-19 Sensitive Element Tracking Deep Audit And Fail-Closed Hardening

### Scope And Checkpoints

- Canonical workspace: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628`.
- Source checkpoints used during this pass include `_backups\element-tracking-sensitive-audit-20260719-110603`, `_backups\element-tracking-sensitive-audit-phase2-20260719-112002`, `_backups\element-tracking-pending-v1-compat-20260719-114657`, `_backups\element-tracking-disable-preservation-20260719-115726`, `_backups\element-tracking-sensitive-audit-phase3-20260719-121317`, `_backups\element-tracking-corrupt-checkpoint-harness-20260719-121805`, `_backups\element-tracking-sensitive-audit-phase4-20260719-122240`, `_backups\element-tracking-external-window-guard-20260719-122756`, `_backups\element-tracking-uncertainty-ui-20260719-122954`, `_backups\element-tracking-time-semantics-20260719-124500`, and `_backups\element-tracking-sensitive-audit-phase5-20260719-131716`.
- QA-script checkpoints: `_backups\family-browser-static-audit-utf8-20260719-123455`, `_backups\family-browser-year-argument-normalization-20260719-123728`, and `_backups\family-browser-ui-harness-utf8-20260719-125442`.
- ProgramData checkpoints before the two final deployments: `_backups\programdata-before-element-tracking-sensitive-final-20260719-131152` and `_backups\programdata-before-element-tracking-sensitive-final5-20260719-133402` (`17` files each, approximately `20 MB`).
- Revit process count remained `0` throughout automated verification and deployment; no user RVT was opened or modified by this audit.

### Sensitive Failure Modes Found And Fixed

- **Incoming activity before a trustworthy session existed:** Synchronize or Reload Latest could begin while the initial baseline capture failed. A later `DocumentChanged` event could then create a late session and misattribute incoming central changes as local work. Per-document synchronize/reload suppression windows now open before baseline/session creation, reject incoming events without relying on a session, and close only at the corresponding completion boundary.
- **A recovered rebase could erase an earlier coverage gap:** `ExternalRebaseFailed` is now sticky until a successful commit/reset baseline. A later successful rebase cannot silently upgrade earlier uncertain attribution to fully trusted history.
- **Checkpoint finalization was not fully fail-closed:** Immutable publication now stops when the local write-ahead checkpoint cannot be finalized. If replay of a finalized checkpoint fails, recovered commits remain attached to the live session for retry instead of being discarded.
- **Corruption could be hidden as a destination mismatch:** Checkpoint load now verifies project identity, envelope integrity, and every signed inner commit before evaluating management-destination binding. A corrupt record cannot be disguised as an ordinary old-folder checkpoint.
- **Corrupt evidence could be counted, deleted, rebound, or overwritten:** pending counts, direct delete, bulk cleanup, destination migration, and normal Save all validate the complete signed payload. Invalid, foreign-project, corrupt, or destination-conflicting checkpoints are preserved for review and cannot be replaced by new evidence.
- **Commit identities were deduplicated too early:** checkpoint input previously grouped commits before assigning missing entry IDs, so two valid records with blank IDs could collapse into one. Every identity and integrity hash is now created first; only byte-equivalent signed duplicates may collapse.
- **Conflicting records could share one entry ID:** synchronization no longer discards duplicate IDs before checkpoint validation. Two records with the same ID but different integrity hashes now fail closed instead of keeping the first and losing the second.
- **An unsigned inner record could hide behind a valid outer envelope:** checkpoint inner commits now require a non-empty, valid integrity hash and empty checkpoints are invalid. Frozen legacy compatibility remains only for immutable history and pending-envelope formats that actually existed before this checkpoint schema.
- **Valid old-destination evidence could be silently rebound:** a normal Save never changes a checkpoint's management destination. Only the explicit source-scoped managed-folder migration can rebind a valid checkpoint; byte identity is preserved before that transition.
- **Schema evolution could invalidate valid old evidence:** immutable integrity v1/v2 and pending-envelope v1 remain readable and replay-idempotent. New records use the frozen integrity-v3 projection and pending-envelope v2. An already signed but invalid commit is rejected, never re-signed as if it were valid.
- **Workshared local Save and central publication time were conflated:** schema 4 now preserves `LocalSaveProtectedAtUtc` separately from `PublishedAtUtc`. Closing/reopening the same local file restores protected edits; only successful synchronization publishes immutable central history.
- **Undo/Redo and rebase ambiguity could guess too much:** unmatched Undo/Redo no longer consumes an arbitrary transaction, active pending IDs are used instead of every historically touched ID, create-then-delete remains explicit transient evidence, and overlap with an external rebase lowers confidence rather than inventing authorship.
- **Central replacement and project aliases could hide history:** stable-path fallback discovers valid history under an obsolete central identity, while checkpoint commits still must match the current project identity. Foreign-project checkpoint records are rejected from publication.
- **Turning tracking OFF could destroy recoverable evidence:** policy disable clears live observation state but does not delete protected local checkpoints. They remain available when tracking is deliberately re-enabled and the project identity is valid.
- **Uncertainty was not prominent in recent history:** the dashboard now counts external-rebase-gap commits and labels them `외부 업데이트 범위 누락 / 검토 필요` rather than presenting them as ordinary client observations.
- **The QA gate itself had two false-failure paths:** Windows PowerShell 5.1 now reads static/UI contracts and result JSON as UTF-8, and build/UI scripts normalize both string-array and comma-delimited year arguments. The Korean unsaved-project fixture therefore no longer becomes mojibake or a false English-language leak.

### Scenario Matrix

| Scenario | Automated Result | Remaining Runtime Boundary |
|---|---:|---|
| Standalone create/modify/delete/create-then-delete reduction | PASS | Confirm native event order in a disposable RVT |
| Matched Undo/Redo and unmatched/grouped ambiguity | PASS | Confirm Revit transaction grouping by version |
| Failed initial baseline during Sync/Reload | PASS, incoming window is independent of session | Force a real API capture failure if practical |
| Successful/failed/cancelled Save, Save As, and Sync | PASS, failed boundaries append no immutable commit | Confirm actual Autodesk status callbacks |
| Workshared local Save, close/reopen, then Sync | PASS, protected and published times remain separate | Confirm same-local-file reopen on a real central model |
| Reload Latest native bridge and 2019 fallback | PASS by source/API binding and lifecycle fixture | Confirm real callback ordering on 2019/2023/2025/2027 |
| Offline write-ahead, reconnect, replay, duplicate retry | PASS and idempotent | Confirm one reachable SMB share |
| TEST-root to homepage-root explicit migration | PASS, source-scoped only | Confirm real managed-root migration permissions |
| Corrupt/tampered envelope, inner commit, collision, or foreign project | PASS, evidence retained and publication blocked | Review operational quarantine/support flow |
| Tracking OFF/ON with protected workshared checkpoint | PASS, checkpoint preserved | Confirm administrator UX in Revit |
| Central file replacement and stable-path history lookup | PASS | Confirm actual central replacement/restore workflow |
| Parallel clients writing separate immutable records | PASS in process-level concurrency fixture | Confirm two physical PCs and mapped-drive/UNC aliases |

### Automated Verification

- Focused tracking audit: `artifacts\family-browser-workflow-audit\element-tracking-sensitive-phase13-20260719\workflow-audit-summary.md`.
- Final complete quality gate: `artifacts\family-browser-ui-audit\20260719-element-tracking-sensitive-final5\quality-gate-summary.md`; every stage is `OK`.
- Workflow audit: `34` workflows, `494/494` checks passed. Tracking workflows include ledger lifecycle, integrity recovery, session recovery, schema compatibility, event ambiguity, policy concurrency, offline recovery, and multi-client retention.
- IE `WebBrowser` harness: `208 OK`, `1` expected Revit 2021 runtime skip, `0` failures; `5,328` clickable controls inspected, `4,176` host-action candidates and `1,072` browser-only clicks executed. The `32` warnings are expected handled empty-selection alerts.
- Performance gate with 1,000 Family plus 1,000 System rows passed. Only `150` DOM rows were rendered per active list; worst shell `11 ms`, cold usable `873 ms`, warm usable `513 ms`, filter `17 ms`, and cache operation `66 ms`.
- Revit 2019/2021/2023/2025/2027 Release build and Stage verification passed. Revit 2021 remains package/Stage/ProgramData verification only because Revit 2021 runtime is not installed on this PC.

### ProgramData Deployment

- Elevated final installation completed with Revit process count `0`. Installed verification reports `OK` for 2019, 2021, 2023, 2025, and 2027.
- Stage-to-ProgramData file counts match exactly and SHA-256 mismatch count is `0` for every target.
- Installed core DLL SHA-256: 2019/2021/2023 `8E2FF7E4F07905E9996E7788B37BD6F2274DF387C92677DD810D1F7E3A314E7C`; 2025 `0A7A0D8015D0C3B119420C193DF407B859DF92F9464A5B1DA7F943B915AE7864`; 2027 `00B8B484EE46160786C0797E0F25D751AC88A2B57F8C85E6BD3CA6DFEE07E2DC`.

### Truth Boundary And Release Conditions

- The automated fixtures prove persistence, reduction, integrity rejection, idempotency, collision behavior, and UI visibility. They do not prove Autodesk's native event ordering on every Revit version; queue item 40 remains the required disposable-model and two-PC sign-off matrix.
- This is instrumented client telemetry, not server-side continuous surveillance. Every editing workstation must run this revision for per-user/per-transaction evidence. An uninstrumented editor creates only a later coverage gap or catalog drift signal.
- Integrity checks detect accidental corruption and inconsistent payloads, but they are not an HMAC or digital signature backed by a private key. A user with write/delete permission to the managed history share can still remove or replace evidence. Production ACLs, backups, retention, and append-only storage policy remain operational requirements.
- Client UTC timestamps cannot prove a conflict-free global order across PCs with skewed clocks. The ledger preserves each client's commit and confidence markers; it does not infer which of two simultaneous edits happened first.
- Element summaries are not a complete parameter-by-parameter, geometry, ownership, deletion, or modeling forensic diff. The feature records only what Revit events and captured before/after summaries expose to the add-in.
- Opening the exact same local RVT in multiple Revit processes, manually replacing/restoring a local file at the same path, and migrating checkpoints while another client writes are not production-proven. Keep these as `Needs Revit Check` or `Needs Design` rather than silently resolving ambiguous evidence.

## 2026-07-19 Sensitive Element Tracking Final Boundary Audit

### Scope And Checkpoints

- This pass continued from the sensitive tracking audit above and re-read the event boundaries, session recovery, checkpoint promotion, shared-history publication, and dashboard truth indicators one more time.
- Source checkpoints: `_backups\element-tracking-local-checkpoint-visibility-20260719-170200`, `_backups\element-tracking-unknown-external-start-20260719-171000`, `_backups\element-tracking-finalized-checkpoint-state-20260719-172000`, and `_backups\element-tracking-commit-protection-visibility-20260719-173000`.
- Audit-document checkpoint: `_backups\family-browser-audit-md-before-sensitive-final6-20260719-174500`.
- ProgramData checkpoint: `_backups\programdata-before-element-tracking-sensitive-final6-20260719-174000` (`17` files, `20,619,233` bytes).
- Revit process count was `0`; this pass did not open or mutate a user RVT.

### Additional Failure Modes Found And Fixed

- **A failed workshared local-Save checkpoint was invisible:** `DocumentSession.LocalSaveCheckpointFailed` now remains sticky until a durable baseline reset. The header and Home board show an unprotected-local-save warning and explicitly tell the user not to close Revit. An empty/no-change Save does not create a false warning.
- **An unreadable Sync or Reload Latest start could leak incoming IDs into local authorship:** workspace-scoped unknown-start suppression now opens even when the start callback cannot expose its `Document`. `DocumentChanged` suppresses incoming IDs before any late session can be created, records an external-rebase coverage gap, and clears the guard only at the matching completion boundary.
- **A finalized checkpoint was described as if another central sync were required:** checkpoint counting now separates ordinary local-Save checkpoints from synchronization-succeeded checkpoints awaiting immutable managed-history promotion. The latter is labeled restart-safe and does not ask for another synchronization.
- **Save/Sync protection failure could look healthy:** `CommitBoundaryProtectionFailed` now covers unreadable completion callbacks, unavailable worksharing state/project identity, checkpoint finalization failure, and standalone immutable-history persistence failure. The warning is cleared only after a genuinely durable boundary succeeds.
- **The dashboard promised Refresh retry but Refresh did not promote finalized checkpoints:** `FlushPending` now enumerates destination-bound finalized checkpoints under the managed file lock, publishes the exact signed commits idempotently, and removes only the exact checkpoint revision through compare-and-swap cleanup. A concurrently newer revision is retained.
- **Checkpoint retry could accidentally imply duplicate publication:** the executable fixture now proves first Refresh promotion, immutable history existence before cleanup, exact-revision cleanup, and a second idempotent Refresh with no duplicate commit.

### Final Automated Verification

- Focused final tracking gate: `artifacts\family-browser-workflow-audit\20260719-sensitive-phase54\workflow-audit-summary.md`; PASS, `34` workflows and `739/739` checks.
- Complete final gate: `artifacts\family-browser-ui-audit\20260719-element-tracking-sensitive-final6\quality-gate-summary.md`. Static/action contracts, lifecycle audit, all five builds and Stage verification, the 2,000-row performance/cache gate, and the IE `WebBrowser` harness passed.
- UI harness: `208 OK`, `1 SKIP runtime-not-installed` for Revit 2021, `0` failures; `5,328` clickable controls, `4,176` host-action candidates, and `1,072` browser-only clicks. The `33` warnings are expected handled empty-selection/fixture alerts.
- Performance remained inside the acceptance limits. Worst measured shell was `37 ms`, cold usable `1,009 ms`, warm usable `573 ms`, filter `18 ms`, and cache operation `85 ms`, with only `150` rows in the active DOM window while every 1,000-row fixture remained reachable.
- The initial integrated install step correctly failed before mutation because the invoking PowerShell was not elevated. The original result is intentionally retained. The same verified installer script was then run in an elevated process and the separate post-install evidence is `artifacts\family-browser-ui-audit\20260719-element-tracking-sensitive-final6\post-install-verification.md`.

### ProgramData Deployment

- Elevated installation completed with exit code `0`, then `Verify-FamilyBrowserRecovered.ps1 -Installed` reported `OK` for Revit 2019, 2021, 2023, 2025, and 2027.
- Stage and ProgramData payload counts match and SHA-256 mismatch count is `0` for every target. All five `.addin` manifests also match Stage.
- Installed DLL SHA-256: Revit 2019/2021/2023 `8796B532F32F1347082F8B9AD2F68C09E3AFD3AAF3A831C583EAEDD5BB025C0C`; Revit 2025 `6D0E5365B1A2F5B1C35753A8488FE4579C16A393C55E17D272DAF543DDB9FD61`; Revit 2027 `241951D3150FCABD9D775B468C4C4CA17C91A513531E48C2BDFF1A76B5208E4E`.

### Residual Truth Boundary

- `Needs Revit Check`: confirm actual Autodesk callback order for successful/cancelled Save, Save As, Sync, Reload Latest, Undo, and Redo in disposable Revit 2019/2023/2025/2027 projects. Revit 2021 remains build/Stage/installed verification only on this PC.
- `Needs Two-PC/SMB Check`: confirm two physical clients, mapped-drive/UNC aliases, simultaneous commits, disconnect/reconnect, and checkpoint promotion against the same real share. The named management-context mutex is machine-local; the managed-file protocol, not that mutex, supplies cross-PC serialization.
- `Needs Design`: a workstation without this add-in cannot provide actor/transaction evidence. Later catalog drift can prove that an unobserved change exists, but cannot recover who changed it or the exact intermediate synchronization in which it occurred.
- `Needs Operations`: SHA-256 integrity detects corruption and inconsistent records but is not a secret-key signature. Managed-share ACLs, backups, retention, and append-only/audit storage remain required against a user who can deliberately delete or replace history files.
- `Needs Design`: client UTC clocks do not establish a perfect global order, Revit username changes are not a stable enterprise identity, and opening the same local RVT in multiple Revit processes remains ambiguous.
- The ledger is not a full BIM forensic diff. It excludes view/internal-only noise by policy and cannot reconstruct every parameter, geometry, ownership, or deleted-element detail that Revit does not expose at the event boundary.

## 2026-07-19 Sensitive Ledger Coverage-Gap And Management-Path Audit

### Scope And Checkpoints

- Canonical workspace: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628`.
- Source checkpoint: `_backups\20260719-sensitive-ledger-gap-v2`.
- ProgramData checkpoint: `_backups\programdata-before-sensitive-ledger-gap-v2-20260719-201649` (`17` files, `21,068,801` bytes).
- This pass re-read the first-event, unreadable-event, Save/Save As/Sync boundary, checkpoint recovery, schema integrity, management-folder transition, history display, and installed payload paths. No user RVT was opened or modified.

### Additional Defects Found And Fixed

- **An uncertainty-only event could disappear:** a late first `DocumentChanged` callback or unreadable element-ID collection incremented uncertainty counters, but `BuildCommit` returned `null` when no trustworthy element row existed. A later successful boundary could then reset the session. Schema 5 now persists a zero-row `CoverageGapOnly` record with explicit `EventReadFailureCount` and `CommitBoundaryReadFailureCount`, without inventing an element ID.
- **Malformed empty evidence could be accepted as success:** non-coverage zero-row commits were silently filtered. A malformed checkpoint update could therefore clear an existing checkpoint. Such input now fails closed; only a coverage-only record with real gap evidence is accepted.
- **New gap fields were not integrity-protected:** schema-5 commits use integrity v4, whose canonical projection includes all coverage fields. A modified gap count fails checksum validation. Legacy integrity v1/v2/v3 history and pending-envelope v1/v2 remain readable and idempotent.
- **Automatic management-path replacement could strand protected local saves:** all three machine-config stores now serialize path changes, reject active uncommitted evidence, reject protected live recovery sessions, and inspect every local checkpoint. Destination mismatch, corruption, read failure, and lock contention all block automatically; only the explicit verified migration workflow can authorize rebinding.
- **Management-context refresh could discard live state:** it now invalidates only the policy cache. Active sessions, workshared recovery state, and Sync/Reload suppression windows survive until a durable boundary.
- **Coverage-only history was visually hidden:** recent-history counts, preview rows, structured HTML, and explicit XLSX now show `관찰 공백 / 요소 ID 확인 불가` records even when there are zero element rows.
- **The workflow contract described an obsolete schema:** the schema compatibility scenario now states schema 5/integrity v4 and the actual legacy compatibility range.

### Final Verification

- Focused lifecycle and persistence audit: `artifacts\family-browser-workflow-audit\20260719-sensitive-ledger-gap-v2-contract-final\workflow-audit-summary.md`; PASS, `34` workflows and `769/769` checks. The executable persistence harness passed `59` checks.
- Core quality gate: `artifacts\family-browser-ui-audit\20260719-sensitive-ledger-deep-audit-core-final5\quality-gate-summary.md`; static/action contracts, lifecycle audit, five-target Stage assembly/verification, and 2,000-row performance/cache are all `OK`.
- Complete IE evidence and consolidated result: `artifacts\family-browser-ui-audit\20260719-sensitive-ledger-deep-audit-final3\deep-audit-completion-summary.md`; `208/208` scenarios, `0` failures, `5,328` clickable controls, `4,176` host actions, and `1,072` browser-only clicks. The `32` warnings are expected handled empty-selection alerts.
- Performance stayed within target: worst shell `14 ms`, cold usable `754 ms`, warm usable `446 ms`, filter `15 ms`, and cache operation `73 ms`; each 1,000-row list rendered only `150` active DOM rows.
- Revit 2019/2021/2023/2025/2027 build, Stage, and ProgramData verification passed. Revit 2021 remains `SKIP runtime-not-installed` for live execution on this PC.

### ProgramData Deployment

- Revit process count before deployment was `0`; elevated installation completed with exit code `0`.
- Stage and ProgramData contain the same `17` payload files. Missing count and SHA-256 mismatch count are both `0`.
- Installed DLL SHA-256: 2019/2021/2023 `48A52CE4A64EDED2DE6DC5D8B494240BAB325546DC4A87BF1BD99BE1D064BA0C`; 2025 `C3A7F4BFA11A7691749767537C22BB9EE47A4569BEEBA950C4CCE40188E4901C`; 2027 `55D7E8AABD5984F8FD2A3D592B48AB27E738BDA104054AB3D6986A0160FA0D3C`.
- Install log: `artifacts\family-browser\programdata-install-20260719-sensitive-ledger-gap-v2.log`.
- Installed verification log: `artifacts\family-browser\programdata-verify-20260719-sensitive-ledger-gap-v2.log`.

### Remaining Truth Boundary

- `Needs Revit Check`: actual Autodesk callback order for Save, Save As, Sync, Reload Latest, Undo/Redo, cancellation/failure, close/reopen, and first-baseline timing on disposable Revit 2019/2023/2025/2027 projects.
- `Needs Two-PC/SMB Check`: simultaneous commits, mapped-drive/UNC/server aliases, disconnect/reconnect, checkpoint promotion, and management-root migration against one real share.
- `Needs Operations`: SHA-256 is consistency/tamper evidence, not secret-key signing. Managed-share ACLs, backups, retention, and append-only storage are still required.
- `Needs Design`: an uninstrumented PC cannot provide actor/transaction attribution, client clocks cannot prove a total global order, and Revit usernames are not immutable enterprise identities.
- Modification events are retained, but arbitrary parameter-by-parameter and geometry deltas are not fully serialized. Views and internal-only Revit elements remain intentionally outside this model-element ledger.

## 2026-07-19 Sensitive Ledger Destination And Semantic Consistency Audit

### Scope And Checkpoints

- Canonical workspace: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628`.
- Source checkpoints: `_backups\20260719-sensitive-ledger-semantic-destination-audit` and `_backups\20260719-sensitive-ledger-cross-record-destination-audit`.
- ProgramData checkpoint: `artifacts\family-browser-ui-audit\20260719-sensitive-ledger-final-backend2\programdata-before-install` (`17` files, `21,099,734` bytes).
- Revit process count before installation was `0`. No user RVT/RFA was opened or changed by this audit.

### Additional Fail-Open Defects Found And Fixed

- **New evidence could be created without a management destination:** element commits and workshared local-save checkpoints now fail before writing when a stable management destination cannot be derived. The same rule now applies to operation audit records and standard-change candidates, including the audit entry required before an administrator changes the element-tracking policy.
- **A destination-less pending element wrapper could later move into an unrelated root:** current pending wrappers require a non-empty destination, entry identity, supported envelope version, and valid checksum. Frozen raw/v1/v2 legacy compatibility remains readable, but a newly shaped wrapper cannot use the legacy empty-destination wildcard.
- **Managed-folder migration could rewrite a checkpoint with an empty target:** target identity is verified before any checkpoint bytes change. Failure increments the explicit migration failure count and preserves the original evidence byte-for-byte.
- **An empty source could authorize an ambiguous migration:** `FlushPendingForManagedFolderTransition` now rejects an empty or unverifiable source before flushing or rebinding anything. A valid transition remains source-scoped.
- **A flush could proceed after managed-root availability and destination identity disagreed:** flushing now fails closed when the current destination identity is empty instead of evaluating pending records against an ambiguous root.
- **A recomputed checksum could legitimize internally contradictory schema-v5 data:** v5 validation now binds the stable project identity to the captured identity/canonical path, verifies tracking/baseline/observation/protection/publication times against the commit boundary, verifies before/after unique identity, and recomputes row-derived counts. Coverage-gap-only records must remain zero-row. Legacy schema/integrity formats retain their frozen compatibility path.

### Automated Verification

- Focused persistence harness: `63/63` checks passed. It now attempts empty operation/candidate/element/checkpoint destinations, recomputed contradictory v5 project/time payloads, empty-source migration, empty-target checkpoint migration, collisions, corruption, replay, concurrent writes, Save As identity changes, and finalized-checkpoint promotion.
- Workflow contract: `artifacts\family-browser-ui-audit\20260719-sensitive-ledger-final-backend2\workflow\workflow-audit-summary.md`; PASS, `34` workflows and `769/769` checks.
- Backend quality gate: `artifacts\family-browser-ui-audit\20260719-sensitive-ledger-final-backend2\quality-gate-summary.md`; static/action contracts, workflow audit, all five builds/Stage verification, and 2,000-row performance/cache are `OK`.
- Full IE `WebBrowser` evidence from the immediately preceding UI-complete gate remains valid because this pass changed only shared persistence and its audit contract: `artifacts\family-browser-ui-audit\20260719-sensitive-ledger-semantic-destination-final2\harness\ui-harness-summary.md`; `208 OK`, one expected Revit 2021 runtime skip, and `0` failures.
- Performance remained within target. Every 1,000-row Family/System fixture exposed all rows through 150-row windows; measured shell, cold/warm usable, filter, and cache times passed the configured limits.

### Build And ProgramData Deployment

- Revit 2019/2021/2023/2025/2027 Release build and Stage verification passed. Revit 2021 remains `SKIP runtime-not-installed` for live execution on this PC.
- Elevated ProgramData installation completed with exit code `0`; independent installed verification reports `OK` for all five targets and exact Stage payload hashes.
- Installed DLL SHA-256: 2019/2021/2023 `CE5C54625EA754ACB9BFAF647E039FE6AA58A695EBBA71C12FF86F6B5E0A8D00`; 2025 `ED3986D240F2A6D457794BC080EBC5F4B9E4A2692CAD118AD40CE6CEE83E6D07`; 2027 `F101D09E2F80FAC29888EC5722FC073A28B01F5ED2B331734FF5EC4FB4812435`.
- Install/status evidence: `artifacts\family-browser-ui-audit\20260719-sensitive-ledger-final-backend2\programdata-install.log` and `programdata-install-status.json`.

### Residual Release Boundary

- `Needs Revit Check` (queue 40): real successful/cancelled/failed Save, Save As, Sync, Reload Latest, Undo/Redo, close/reopen, late baseline, and large-model timing on disposable projects.
- `Needs Two-PC/SMB Check` (queue 40): mapped-drive/UNC aliases, two physical clients, simultaneous commits, disconnect/reconnect, checkpoint promotion, and TEST-to-homepage migration against one authoritative share.
- `Needs Security/Operations` (queue 42): SHA-256 here is unkeyed consistency evidence, not an HMAC or digital signature. It cannot prove who wrote a plausible replacement, and without an anchored manifest it cannot detect deletion of an entire valid record/day folder. Operation and standard-candidate records currently have no record-level checksum.
- `Needs Design`: uninstrumented workstations cannot provide actor/transaction attribution; client clocks cannot establish a total global order; Revit usernames are not immutable enterprise identities.
- The ledger records model-element IDs and selected before/after summary metadata. It is not a complete parameter, geometry, ownership, or journal-level forensic reconstruction.

## 2026-07-19 Sensitive Tracking State, Storage, And Catalog Portability Audit

### Scope And Backups

- Canonical workspace: `C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628`.
- Source checkpoints: `_backups\20260719-sensitive-external-gap-policy-boundary-audit`, `_backups\20260719-sensitive-history-attribution-visibility`, `_backups\20260719-sensitive-enumeration-failclosed`, and `_backups\20260719-sensitive-project-catalog-portability`.
- ProgramData checkpoint: `_backups\programdata-before-sensitive-tracking-final-review-20260719-225900` (`17` files, `21,128,646` bytes).
- Revit process count before installation was `0`. No user RVT or RFA was opened or modified during this pass.

### Additional Defects Found And Fixed

- **A remote tracking OFF could discard already observed evidence or start a foreign document session from stale policy state:** policy/session decisions are now explicit. Existing uncommitted evidence survives only to the next successful Save/Sync boundary, protected recovery remains recovery-only, and a different document cannot start tracking from fallback/deferred state.
- **External update rebase failure could disappear when no trustworthy element row existed:** schema-v5 zero-row coverage evidence now survives Save/Sync. The history summary and attribution column show external rebase gaps even when an incoming-overlap warning also exists.
- **Independent uncertainty reasons overwrote each other:** incoming overlap, external rebase gap, Save-boundary gap, DocumentChanged gap, Undo/Redo ambiguity, missing Revit user, and missing element ID are now composed as separate review labels instead of selecting only the first reason.
- **Unreadable local queue/checkpoint/history folders looked like zero records:** enumeration failure now fails closed. Pending count stays nonzero/unknown, checkpoint status reports lock/read unavailability, managed-folder switching and rebinding are blocked, flush reports failure, and immutable-history loading increments invalid evidence rather than showing a clean empty history.
- **Project name-catalog snapshots were bound to one PC's absolute mapped-drive/UNC spelling:** new manifest references are relative to the project catalog folder. Legacy absolute references resolve only by their snapshot filename under the current managed project folder, preventing traversal into an old or unrelated management root.
- **Central-file replacement could orphan the lightweight project catalog:** project catalog folders now key from stable project path identity. Existing file-identity folders remain discoverable by their manifest; multiple legacy candidates fail closed instead of selecting an arbitrary accepted baseline.
- **Corrupt catalog state could be silently replaced as `baseline missing`:** unreadable manifests/states, manifest-state mismatches, missing referenced snapshots, duplicate/tampered entry keys, hash mismatch, counter mismatch, unsupported schema, and foreign project identity now produce an error while preserving the existing files.
- **A matching Browser operation was presented too strongly as proof of the user:** the UI and XLSX now say `Browser 작업 기록 일치 / 작업자 미확정`. Name/time matching helps triage but is not proof of the actor or exact synchronization.
- **The 200-commit/5,000-row history window could hide older evidence without saying so:** the loader reports total valid commits before applying its display limit, and the result becomes a warning when either limit truncates the visible/exported view.

### Automated Verification

- Integrated quality gate: `artifacts\family-browser-ui-audit\20260719-sensitive-tracking-final-review\quality-gate-summary.md`; all six stages are `OK`.
- Workflow lifecycle audit: `35` workflows, `812/812` checks passed. Executable persistence harness: `70/70` checks passed, including corrupt/enumeration failures, policy boundaries, collisions, replay, path replacement, local Save checkpoints, and multi-writer retention.
- IE `WebBrowser` harness: `208 OK`, `1 SKIP runtime-not-installed` for Revit 2021, `0` failures. It inspected `5,328` clickable controls, executed `1,072` browser-only clicks and `4,176` host-action candidates. The `33` warnings are handled empty-selection/fixture alerts.
- Performance gate passed with 1,000 Family and 1,000 System rows per target. Worst shell was `12 ms`, cold usable `823 ms`, warm usable `451 ms`, filter `14 ms`, and cache operation `70 ms`; only `150` active DOM rows were rendered while all 1,000 remained reachable.
- Revit 2019/2021/2023/2025/2027 Release build and Stage verification passed. Revit 2021 remains build/Stage/ProgramData verification only because its runtime is not installed on this PC.

### ProgramData Deployment

- Elevated installation and installed verification passed for 2019, 2021, 2023, 2025, and 2027. Stage-to-ProgramData SHA-256 mismatch count is `0` for every target.
- Installed DLL SHA-256: 2019/2021/2023 `3250C5AB9941548102EDE34C1EF3613EA6838C5FA76E493B9165DE3572C6D9E5`; 2025 `C80F0B1AB8A99CFDBA669A0A41940C7BFF8F866EB895922119FD65E8106F4CEF`; 2027 `55E76280E26FC68825F9296286F271DEDFB9B69F61993E5BCF5B42053DD686D3`.
- Install evidence: `artifacts\family-browser-ui-audit\20260719-sensitive-tracking-final-review\programdata-install.log` and `programdata-install-status.json`.

### Release Truth Boundary

- `Needs Revit Check`: use disposable projects to confirm actual Autodesk callback order for successful/cancelled/failed Save, Save As, Sync, Reload Latest, Undo/Redo, workshared close/reopen, remote policy OFF, and late first-event capture in Revit 2019/2023/2025/2027.
- `Needs Two-PC/SMB Check`: confirm two physical clients, mapped-drive/UNC/DNS/IP aliases, simultaneous commits, disconnect/reconnect, local-checkpoint promotion, central replacement, and management-folder migration against one real share.
- `Needs Security/Operations`: SHA-256 is an unkeyed consistency check, not HMAC/digital signing. A writer with sufficient share permission can recompute hashes, and deletion of a complete valid record/day folder cannot be detected without an externally anchored append-only manifest, WORM store, server audit, ACLs, and backups.
- `Needs Security/Operations`: operation logs and standard-change-candidate history do not yet have record-level integrity; only their pending envelopes are checksummed. Do not treat name-only Browser operation matching as forensic actor proof.
- `Needs Design`: an uninstrumented workstation can be detected later only through name-catalog drift. It cannot reveal the exact user, intermediate synchronization, parameter/geometry edit, or deletion moment that was never observed by the add-in.
- `Needs Design`: client clocks do not establish a total cross-PC order, Revit usernames are mutable identity labels, and there is no retention/archive policy for indefinitely growing immutable history.
- A crash before a successful standalone Save can lose unsaved in-memory telemetry. Workshared local Save receives restart-safe checkpoint protection, but this still requires real Revit fixture confirmation.
- Tracking intentionally excludes Views, DataStorage, ProjectInfo, negative temporary IDs, and categoryless Revit internals. Parameter-only/geometry changes are retained as generic `Modified` when DocumentChanged supplies the element ID; the ledger is not a complete parameter-by-parameter or geometry forensic diff.

## 2026-07-20 Family Browser Support Links And Product Update Check

Status: `Complete`

### Backup

- Source checkpoint: `_backups\20260720-family-browser-support-update-links`
- Live-link correction checkpoint: `_backups\20260720-family-browser-live-support-links-fix`
- Canonical `1.0` correction checkpoint: `_backups\20260720-family-browser-version-1.0-correction`

### Implemented

- Added an always-visible `Support / 지원` sidebar group with `Check for Updates / 업데이트 확인`, `Homepage / 홈페이지`, and `Manual / 매뉴얼` entries for both administrator and modeler modes.
- Added shared `FamilyBrowserProductUpdateService` code for all Revit hosts. A manual check reads the live `https://update.zerokky.com/Release/family-browser/latest.json` feed on a background task, bypasses stale HTTP cache, compares semantic versions against the canonical Family Browser version `1.0`, and presents current/latest version, publication date, and release notes in the common HTML result dialog.
- The update action changes only the status text while the request is running, rejects duplicate concurrent clicks, logs success/failure diagnostics, and offers the product homepage when a newer version exists. It does not run through a Revit mutation external event.
- The Homepage action validates an HTTP/HTTPS constant before opening it with the Windows default browser.
- The visible Manual action now opens the official Family Browser manual at `https://update.zerokky.com/family-browser/index.html` in the default browser. It no longer selects the internal Help pane.
- Added a modeler-English manual render scenario and required `update-check` / `open-homepage` / `open-manual` routes to the UI contract.
- Removed the incorrectly added local `Sever\family-browser\latest.json` source and README references; that path is not the production release feed.

### Deployment Verification

- ProgramData backup is `_backups\programdata-before-family-browser-support-links-20260720-095030` (`17` files, `21,203,430` bytes).
- Pre-correction-install backup is `_backups\programdata-before-live-support-links-fix-20260720-101148` (`17` files, `21,294,514` bytes).
- Canonical-version install backup is `_backups\programdata-before-family-browser-version-1.0-correction-20260720-103330` (`17` files, `21,297,822` bytes).
- Revit process count was rechecked as `0`. Elevated ProgramData deployment completed with exit code `0`, and independent installed verification reports `OK` for 2019/2021/2023/2025/2027.
- Final Stage-to-ProgramData DLL hashes match for every target: 2019/2021/2023 `19DA80F54FEA8F3B19A1F25E7E7F29EBE06FBF790C37B978FB8E467010B4DBF0`; 2025 `BBDE5600F1462E36F5692D0AD69829A716EAAA26CCC70B2141405852EDF796F8`; 2027 `C4B77261545CA19B166BEA6B1E3E0AFAE2FF3618E49E5ACC511A5E2CDABCBABD`.
- Final live endpoint check: release feed HTTP `200` with version `1.0`; official `v1.0` installer is `3,781,529` bytes with SHA-256 `1435E8E06F0DDF0430D4117BC5E2B8E6AE312DBB50B88426AEB78D02516C6BF5`; the incorrectly numbered patch installer URL now returns HTTP `404`.

### Verification Finding

- User runtime evidence exposed two incorrect assumptions in the first implementation: the Manual menu targeted the internal Help pane, and the update URL targeted a nonexistent `/family-browser/latest.json` endpoint. The production homepage identifies `/family-browser/index.html` as the Family Browser manual and `/Release/family-browser/latest.json` as the product feed; both are now contract constants.
- The new modeler-English manual scenario exposed a real language-refresh defect: the health value initialized as standalone Korean `준비` was not covered by the note translation map and remained visible after switching to English.
- Added exact, symmetric `Ready` / `준비` normalization in all three dashboard hosts. The mapping is intentionally exact so Korean phrases containing `준비` are not partially converted into mixed-language text.
- Hardened product version comparison so semantically equal `1.0`, `1.0.0`, and `1.0.0.0` values normalize to the same four-part version. A nonnumeric remote version now fails visibly instead of falling back to lexical comparison or reporting a false update.

### Automated Verification

- Post-correction static action/route and contract checks passed in all three source hosts: `85` generated actions, `264` exact routes, `64` prefix routes, and `11` browser functions per host.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage integrity verification passed. Existing compiler warning families remain; this change introduced no build error.
- Final Stage DLL SHA-256 after canonical `1.0` correction: 2019/2021/2023 `19DA80F54FEA8F3B19A1F25E7E7F29EBE06FBF790C37B978FB8E467010B4DBF0`; 2025 `BBDE5600F1462E36F5692D0AD69829A716EAAA26CCC70B2141405852EDF796F8`; 2027 `C4B77261545CA19B166BEA6B1E3E0AFAE2FF3618E49E5ACC511A5E2CDABCBABD`.
- Live-link focused IE `WebBrowser` regression: `artifacts\family-browser-ui-audit\20260720-live-support-links-fix`; `24 OK`, one expected Revit 2021 `SKIP runtime-not-installed`, and `0` failures. The Support group exposes the external-manual action in both Korean and English for every runnable host.
- Focused post-fix IE `WebBrowser` regression: `artifacts\family-browser-ui-audit\20260720-support-update-links-manual-regression`; `24 OK`, one expected Revit 2021 `SKIP runtime-not-installed`, and `0` failures. Both Korean and English manual renders passed for every runnable host.
- Full post-fix Revit 2027 UI regression: `artifacts\family-browser-ui-audit\20260720-support-update-links-2027-final`; all `54` scenarios passed, covering home, family/system lists, manual, messages, detail, admin, language, click routing, and 1280/1600/1920 layout cases.
- The earlier integrated gate completed static contracts, nested-family checks, system-apply guards, durable-tracking workflows, all five builds/Stage checks, 2,000-row performance checks, and complete 2019/2023/2025 UI runs before the parent shell reached its 15-minute limit. Its only completed-host failure was the standalone `준비` English-manual finding above; the focused post-fix run closes that regression. The separate full 2027 run closes the seven scenarios left unfinished by that shell timeout.

### Canonical Version Correction

- Product version policy is `1.0`; no patch suffix is used for the current Family Browser release. `FamilyBrowserProductUpdateService.CurrentProductVersion`, its static contract, the public feed, release history, installer filename, and Inno Setup `AppVersion` now all use `1.0`.
- Rebuilt the official five-target installer from the corrected source: `KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_official-version-correction-20260720_Setup.exe`; build metadata records version `1.0`, `17` Stage payload files, `3,781,529` bytes, and SHA-256 `1435E8E06F0DDF0430D4117BC5E2B8E6AE312DBB50B88426AEB78D02516C6BF5`.
- Published the same bytes under the stable official URL `https://update.zerokky.com/Release/family-browser/official/KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_Setup.exe`. Release repository commit: `2888d651541c5cdcd864e0dd2269e948ab721c2f`.
- GitHub Pages deployment run `29710997128` completed successfully for that commit. Cache-bypassed public checks confirmed feed version `1.0`, exact installer hash equality, and removal of the incorrectly numbered patch installer.
- Post-correction static route/contract verification passed; all five Release builds and Stage verification passed; elevated ProgramData installation passed with `0` missing and `0` mismatched files. Revit 2021 remains `SKIP runtime-not-installed` for live execution on this PC.

## 2026-07-20 External Purchase Feature Brief

Status: `Complete`

### Scope And Evidence

- Prepared a source-grounded Korean product brief for an external company purchase and technical review. The canonical version is `1.0` and the document distinguishes implemented behavior, automated evidence, conditional features, and items that still require customer Revit/SMB acceptance testing.
- Evidence was traced from the UI action/visibility contract, workflow lifecycle contract, current dashboard hosts and shared services, this audit ledger, the 2026-07-19 integrated quality gate, performance/cache reports, sensitive tracking reports, and the 2026-07-20 support/update regression.
- The document covers the complete operational surface: dashboard, standard RVT and approved-list administration, precision scan, Family/System lists and details, fingerprint/CSV/nested-family/3D data, load/apply and Save/Sync lifecycle, model check, requests, file Guard, managed-folder bootstrap/migration, standard revision tracking, project name catalog, opt-in element ledger, performance architecture, Korean/English UI, support links, release compatibility, limitations, and buyer acceptance tests.

### Deliverables

- Markdown source: `docs/commercial/KKY_FamilyBrowser_1.0_기능_상세설명서_20260720.md`.
- Editable Word document: `docs/commercial/KKY_FamilyBrowser_1.0_기능_상세설명서_20260720.docx`.
- Submission PDF: `docs/commercial/KKY_FamilyBrowser_1.0_기능_상세설명서_20260720.pdf`.
- Reproducible document generator: `tools/docgen/generate_family_browser_commercial_doc.py`.
- The final PDF is A4, `40` pages, and contains the current product icon plus five current UI/detail screenshots from the automated audit fixtures.

### Document QA

- Rendered and visually reviewed all `40` PDF pages. Cover, contents, headings, tables, screenshots, Korean glyphs, margins, headers/footers, and page boundaries have no clipping, overlap, missing media, or broken layout.
- Corrected the DOCX generator so each independent ordered list restarts at `1` rather than continuing numbering from a previous chapter; the affected quality-gate and conclusion pages were rerendered and checked individually.
- Final text/package checks found `0` local `C:\Users` paths, `0` TODO/MISSING/image-missing markers, `0` replacement characters, and `0` obsolete patch-version strings. PDF metadata title is `KKY Family Browser 1.0 기능 상세 설명서` and page size is A4.
- The DOCX package contains six embedded images and valid A4 section geometry. The PDF contains `40` readable pages and no executable JavaScript.

### Change Boundary

- No Family Browser product source, Revit payload, installer, or ProgramData file was changed for this documentation task. ProgramData was therefore intentionally left unchanged.
- The only code change is the standalone commercial-document generator numbering fix; it is outside the shipped Revit add-in.

## 2026-07-20 Project Management Feature Guide Reframing

Status: `Complete`

### Purpose And Backup

- Reframed the initial technical feature brief as a plain-language `프로젝트 관리 기능 안내서` for BIM managers, designers, reviewers, and project managers.
- Preserved the previous Markdown, DOCX, PDF, and generator in `_backups\20260720-family-browser-project-management-guide-rewrite` before editing.
- The current deliverables replace the earlier technical-review wording while keeping the same stable output paths under `docs/commercial`.

### Content Changes

- Removed purchase wording, Revit target-version matrices, build/Stage/ProgramData details, automated harness counts, internal schema/cache names, and development-machine caveats from the reader-facing document.
- Reorganized the guide around user outcomes: standard consistency, faster library discovery, pre-load comparison, request handling, file-level Guard, managed-folder collaboration, standard revision awareness, project-name drift detection, Save/Sync confirmation, and record-based issue triage.
- Added plain-language explanations of how retained records help narrow a problem period and affected scope, while preserving the truth boundary for uninstrumented PCs, observation gaps, and non-forensic element tracking.
- Added role-based use, six real project scenarios, recommended operating cadence, and a complete end-user feature summary.

### Document Design

- Kept the `standard_business_brief` structure with A4 geometry, but changed typography to ordinary Arial for Latin text and Malgun Gothic for Hangul. Body text is explicitly Regular with restrained dark-gray ink; headings use quiet navy rather than bright decorative styling.
- Replaced the cover subtitle and metadata with document purpose, management scope, and primary users. Running headers now say `프로젝트 표준·라이브러리 관리 기능 안내서`.
- Removed body page-break directives that created large lower-page gaps. The guide now flows naturally in `24` pages instead of the earlier `40`-page technical report.

### Final QA

- Visually inspected all `24` rendered A4 pages at `1191 x 1684` pixels. Cover, contents, headings, tables, lists, five UI/detail screenshots, Korean/English text, headers, footers, and page breaks have no clipping, overlap, missing glyphs, or broken table geometry.
- Final Markdown, DOCX, and PDF each contain `0` occurrences of purchase/development wording, build/Stage/ProgramData terms, target-version matrix years, `1.0.1`, local `C:\Users` paths, TODO/MISSING/image-missing markers, replacement characters, or the temporary LibreOffice install token.
- DOCX is valid A4 with six embedded images and explicit Arial/Malgun Gothic style references. PDF metadata title is `KKY Family Browser 1.0 프로젝트 관리 기능 안내서`, contains `24` pages, and contains no PDF JavaScript.
- Final SHA-256: Markdown `A447A9CE72CD48DEFFA86835064114082371E26602E58797C4B8202C3E93F721`; DOCX `8F3E9FE67B4E7F2200B2C226F82F08EBCE619B5A89C24E772884069A7445734B`; PDF `388C2A0803D1DC5AB4B664C7E29367C7B0B70BB821121130A0A38782250DFE29`.

### Change Boundary

- No Family Browser add-in source, Revit payload, installer, or ProgramData content changed. Only the standalone guide source and its document generator were revised, so ProgramData deployment was intentionally not repeated.

## 2026-07-20 Gulim Typography And Runtime-Test Disclosure Revision

Status: `Complete`

### Backup And Scope

- Preserved the pre-change Markdown, DOCX, PDF, and generator in `_backups\20260720-family-browser-gulim-runtime-check-pagination`.
- This revision changes only the project-management guide and its standalone generator. Family Browser product code, Revit payloads, installers, and ProgramData were not changed.

### Typography And Page Flow

- Changed all reader-visible document text to `Gulim`, including the cover, body, headings, tables, captions, headers, footers, and inline emphasis. The final `word/document.xml` contains only `Gulim` run-font declarations; no Arial, Malgun Gothic, Consolas, or Courier text run remains.
- Added widow/orphan protection and grouping for headings with their introductory paragraphs, images with captions, short ordered and unordered lists, and each `Situation / Use / Result` scenario block.
- Added an intentional page break before the final conclusion so the heading, numbered outcomes, and closing paragraph remain together instead of leaving a single closing paragraph on a mostly blank page.

### Runtime-Test Disclosure

- Added an early reader-facing explanation that implemented behavior is distinguished from behavior that still needs confirmation in a real Revit, project-file, worksharing, permission, and network environment.
- Added focused `실제 환경 확인 필요` callouts for precision-scan warning handling, Family load/update Save-Sync lifecycle, System Type application, file Guard, managed-folder migration/connectivity, and concurrent change/element tracking.
- Added Section `19.1 실제 환경 확인이 필요한 기능` with six rows explicitly marked `기능 반영, 실제 테스트 필요` and concrete acceptance checks for representative RFA/RVT, routing/layer/material systems, local/central and mapped/UNC paths, two-user worksharing, and network interruption/recovery.

### Final QA

- Regenerated the canonical DOCX and PDF and visually inspected all `31` rendered A4 pages at `1191 x 1684` pixels. No clipping, overlap, broken table geometry, missing media, orphaned conclusion text, or missing Korean glyph was found.
- DOCX package verification: A4 section size `11906 x 16838` DXA, six embedded images, and only `Gulim` in reader-visible document runs. Symbol remains only as the list-bullet glyph and an unused built-in Courier style remains in the template style table; neither is used for document text.
- PDF verification: `31` pages, metadata title `KKY Family Browser 1.0 프로젝트 관리 기능 안내서`, six `기능 반영, 실제 테스트 필요` status entries, and no PDF JavaScript.
- Content hygiene checks found zero purchase/development/build-target wording, `1.0.1`, local `C:\Users` paths, TODO/MISSING markers, replacement characters, or temporary install tokens.
- Final SHA-256: Markdown `DB90025AD2BAA547871195129B50AE9A1FF015ABE6626BDAC2A4BD84C0A55F92`; DOCX `58DEA3115D3F99C39183C32E6E798B098BDF2F646CC3AA98906562D84B6273E6`; PDF `49CA3FB8CDEDDE631EDF375559DE525099FD0537F0B78011FC5DE2F96C16E7AF`.

## 2026-07-20 Admin Profile / Admin Mode OFF Effective Guard Correction

Status: `Complete / Needs Revit Check`

### Intended Boundary

- An administrator profile with `Admin Mode OFF` now runs as the effective modeler for guarded project operations.
- For an RVT explicitly registered in File Guard with `BlockFamilyLoadAndEdit`, the Family Browser load command and Revit native Family load/edit routes are blocked while Admin Mode is OFF.
- For a guarded RVT with `BlockTypeChanges`, Family/Type rename and type add/delete routes are blocked while Admin Mode is OFF.
- `Admin Mode ON` restores administrator capability immediately. RVTs that are not File Guard targets remain unaffected; this is not a global restriction on every project.

### Defects Found And Fixed

- `FamilyBrowserNativeCommandGuardService.LoadAdminModeEnabledSetting()` treated every administrator profile as ON and rewrote the saved setting to `true`. It now restores the persisted local ON/OFF setting and returns OFF for a saved administrator-profile OFF state.
- Dashboard permissions used the profile-only `Can(...)` path, so an administrator profile could retain an active `선택 항목 로드` button while Admin Mode was OFF. Dashboard action permissions and status text now use `CanNativeGuard(..., _adminModeEnabled)`.
- The dashboard permission snapshot cache did not include Admin Mode and was not invalidated on mode changes. The cache key now includes `admin-mode=on/off`, and `SetAdminMode` clears the cached snapshot before the immediate shell refresh.
- File Guard already carried a `LoadFamilies` branch but rejected that permission before target evaluation. `ResolveFileGuardPermission` now evaluates `LoadFamilies`, so the Browser load button follows the same guarded-file boundary as native Family editing.
- Added a contract scenario for `administrator profile + Admin Mode OFF + guarded RVT` and a real IE `WebBrowser` assertion requiring zero active Family-load links and exactly one disabled Family-load control.
- The ProgramData installer now permits a non-elevated process only when every exact target Addins folder passes an explicit create/delete write probe. Existing installed DLL ACLs still required the normal elevated replacement path on this PC.

### Backups

- Source checkpoint: `_backups\admin-off-effective-guard-20260720-2015`.
- ProgramData checkpoint: `_backups\programdata-family-browser-admin-off-20260720-202618`.
- Revit process count before deployment: `0`.

### Automated Verification

- Static action/permission contract passed for all three source hosts: `85` generated actions, `264` exact routes, `64` prefix routes, and `11` browser functions per host.
- Workflow audit passed at `artifacts\family-browser-workflow-audit\20260720-202018`.
- Focused IE `WebBrowser` regression passed at `artifacts\family-browser-ui-audit\20260720-admin-off-effective-guard-rerun`: Korean and English protected-file scenarios passed for Revit 2019, 2023, 2025, and 2027 with `0` failures. Revit 2021 is `SKIP runtime-not-installed`.
- The policy seam additionally verified Admin OFF blocks `LoadFamilies`, `EditFamilies`, `AddDeleteTypes`, and combined Family/Type rename on a matching target; Admin ON allows them; a non-target RVT remains allowed.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage verification passed.

### ProgramData Deployment

- Elevated installation completed with exit code `0`; independent installed verification reports `OK` for 2019/2021/2023/2025/2027.
- Stage and installed DLL hashes match exactly: 2019/2021/2023 `630963F41C2DF89E78F24D4E14379861F46452D40B519B9128A9DF6692C9C3AE`; 2025 `CD9D4115BC9D2A826E808B77FF1FC8A5BF71EF3D022933F686BACC7A1C31E6B1`; 2027 `4B23F89FFBEB281A67ECDA9BC5399E8FE89A280E24524EF933FC10138103DEA3`.

### Runtime Acceptance Check

- `Needs Revit Check`: open a disposable RVT that is registered in File Guard, sign in with an administrator profile, switch Admin Mode OFF, and verify Browser `선택 항목 로드`, native `Load Family`, Family Edit, Family/Type rename, and type add/delete are blocked. Switch ON and verify the same actions are restored without pressing Refresh.

## 2026-07-20 Admin OFF Startup State Override Regression

Status: `Complete / Needs Revit Check`

### Runtime Evidence And Root Cause

- The user reproduced an administrator-profile OFF failure with `ProjectKKK.rvt`, whose File Guard record correctly enabled Family load/edit, Family/Type rename, and nested-only standalone placement blocking.
- The loaded Revit 2019 assembly was confirmed to be the then-current ProgramData build, and the managed `standard-policy.json` contained the correct exact RVT path and enabled block flags. This was not a stale-install or policy-save failure.
- A second dashboard lifecycle path remained after the earlier native-guard correction: `CompleteInitialOpenRefresh` assigned `_adminModeEnabled = CanEnableAdminMode(_standardPolicy)` and immediately saved it. It therefore treated administrator capability as the selected Admin ON state and overwrote an explicit OFF choice.
- Manual TEST management-folder configuration repeated the same assignment. Multiple Revit processes could consequently race through dashboard startup and rewrite `%LOCALAPPDATA%\KKY\FamilyBrowser\Settings\admin-mode.txt` to `on`; the reproduced file was last written as `on` at 20:38:17.

### Correction

- Added `ApplyAdminModeAfterPolicyLoad(...)` and the pure `ResolveEffectiveAdminMode(...)` seam to all three dashboard hosts.
- Effective mode is now `requested ON && administrator capability`. Administrator capability alone can never turn the mode on.
- Initial preload restores the persisted user selection; management-folder policy changes preserve the current session selection. Either path may revoke an invalid ON selection to OFF, but neither path can create a new ON selection.
- Both policy-load paths clear the dashboard permission snapshot and push the exact effective state plus already-loaded policy to `FamilyBrowserNativeCommandGuardService`.
- Added runtime diagnostics recording requested, allowed, effective, and persisted/session source values without exposing policy contents.
- Restored the user setting to `off` after deployment because the old code had overwritten the explicit OFF choice.

### Regression Coverage

- Static UI/action contract passed for all three source hosts: `85` generated actions, `264` exact routes, `64` prefix routes, and `11` browser functions per host.
- Workflow audit passed at `artifacts\family-browser-workflow-audit\20260720-205058`: `820/820` checks across `35` workflows.
- The real assembly harness now invokes the Admin selection/capability truth table directly: `OFF + capable = OFF`, `ON + capable = ON`, and `ON + not capable = OFF`.
- Focused IE `WebBrowser` regression passed at `artifacts\family-browser-ui-audit\20260720-admin-off-startup-state-fix` for Korean and English protected-file scenarios on Revit 2019, 2023, 2025, and 2027 with zero failures. Revit 2021 remains `SKIP runtime-not-installed` on this PC.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage verification passed.

### Installer Preflight Correction

- The non-elevated ProgramData write probe could create a new file in each Addins root but could not replace ACL-protected existing DLLs. This allowed a replacement attempt to reach the first delete operation before access was denied.
- A non-elevated replacement now fails before any installed payload mutation whenever an existing Family Browser payload or manifest is present. Fresh installation still requires every exact Addins root to pass the create/delete probe.
- The elevated five-version installation completed with exit code `0`; independent installed verification passed for every target.

### ProgramData Deployment

- Source checkpoint: `_backups\admin-off-startup-state-fix-20260720-204704`.
- ProgramData checkpoint: `_backups\programdata-admin-off-startup-state-fix-20260720-205422`.
- Stage and installed DLL SHA-256 match exactly: 2019/2021/2023 `3A8299948C58ED9421C5D811795614A6DBD862642282611EC68F7A5240BAB649`; 2025 `6E445E2AA054B001F7092F15790CD62EAB967FC70E383178BDFABC5A7C9EE32E`; 2027 `6B0A61E6B6BE82C76275210692BF0CE6B23A741B070078B344A255273736EC42`.

### Runtime Acceptance Check

- `Needs Revit Check`: relaunch Revit 2019 with `ProjectKKK.rvt`. The browser must start with Admin Mode OFF, and the registered RVT must disable Browser Family load plus Revit native `Load Family`, Family Edit, Family/Type rename, and type add/delete. Turning Admin Mode ON must restore the same commands immediately without Refresh; turning it OFF again must remain OFF after browser refresh and reopen.

## 2026-07-20 Admin OFF Native Revit Command Enforcement Follow-up

Status: `Complete / Needs User Revit Check`

### Confirmed Root Causes

- Runtime diagnostics from the protected `ProjectKKK.rvt` proved the file-policy decision itself was correct: Admin Mode OFF, target matched, and effective Family load/edit/type permissions were all `false`.
- The protected-change updater had previously been registered from a modeless dashboard callback. Revit rejected that call as an inactive external application, leaving Family/Type modifications without a transaction-level rollback guard. Registration now occurs once in `OnStartup`, where `UpdaterRegistry` is valid.
- Revit's built-in `Load Family` ribbon state was not refreshed by the permission result. `CanExecute` and pre-document cancellation remain enforcement layers, and the visible Autodesk ribbon control is now synchronized directly on every Admin ON/OFF transition.
- Some Revit releases may not expose a bindable `Load Family` command ID. The ribbon synchronization previously skipped all work when that binding was absent. It now falls back to the always-registered `FamilyLoadingIntoDocument` permission definition, so missing command binding cannot fail open.
- Full recursive traversal of the Autodesk ribbon made Admin switching slow. The implementation now discovers only the `Load Family` ribbon item path, caches the matched controls, and updates their enabled state directly.
- The old single Admin button mixed current state and requested action. The header now exposes explicit ON and OFF choices, and the first eligible administrator run starts ON while subsequent explicit choices remain persisted.

### Enforcement Layers

- Browser load/apply controls use the same effective Admin Mode and File Guard decision as the native guard.
- Native `Load Family` is guarded by command availability, pre-execution cancellation when a binding exists, the cancellable `FamilyLoadingIntoDocument` event, and direct ribbon-state synchronization.
- Family/Type rename and type add/delete are guarded by startup-registered Dynamic Model Updater triggers on `Family`, `FamilySymbol`, and `ElementType`; prohibited transactions receive a blocking Revit failure before commit.
- Admin Mode changes clear both dashboard and native decision caches immediately. No Refresh click is required.

### Automated Verification And Deployment

- Revit 2019/2021/2023/2025/2027 Release builds completed without compile errors.
- Static UI/action contract and native-guard checks passed; all three source hosts expose `85` generated actions, `264` exact routes, `65` prefix routes, and `11` browser functions.
- Workflow audit passed at `artifacts\family-browser-workflow-audit\20260720-223754`.
- Focused IE `WebBrowser` protected-file regression passed at `artifacts\family-browser-ui-audit\20260720-223100` for Korean and English on 2019/2023/2025/2027; 2021 remains `SKIP runtime-not-installed`.
- Installed verification reports `OK` for 2019/2021/2023/2025/2027. Every staged payload file matches the corresponding ProgramData SHA-256.
- Final host DLL SHA-256: 2019/2021/2023 `E27976664199B8E5C087C27B3A807CB595BB4ADC7EE91C8010F20861F923C66D`; 2025 `A387185A0F536713BBE79B25EF6E3EA77A4AD4AEF1A38A5747010AC149FFB46A`; 2027 `50AC41781DDBAF6F4AA6A4BE08B42E98E30CFDE526514EF2DEE6842B098336DC`.

### Runtime Acceptance Boundary

- `Needs User Revit Check`: automatic Revit launch was stopped because it can trigger a license error on this PC. No further automatic Revit launch will be performed.
- After fully restarting Revit, open the registered target RVT and select Admin OFF. Confirm the native `Load Family` button becomes disabled immediately, Browser Family load is disabled, Family/Type rename cannot commit, and type add/delete cannot commit. Then select Admin ON and confirm those commands return without Refresh.

## 2026-07-20 Repeated Project Browser Type Rename Guard Correction

Status: `Complete / Needs User Revit Check`

### Runtime Evidence And Root Cause

- The protected-change security audit recorded only the first rename attempt, while the Revit journal recorded later Family/System Type rename transactions as successful. The effective File Guard policy and Admin OFF state were therefore correct, but repeated native rename commands were escaping the command-level guard.
- The journal identified the actual Project Browser F2/rename command as `ID_PRJBROWSER_RENAME`. The native command definition only listed generic rename IDs such as `ID_RENAME`, so it did not bind the command Revit was actually executing.
- The first attempt happened to be caught by the Dynamic Model Updater fallback. That fallback is a transaction safety layer, not a reliable replacement for cancelling every Project Browser rename before execution, which explains why the first attempt showed an error and subsequent attempts could commit.

### Correction

- Added `ID_PRJBROWSER_RENAME` as the primary Family/Type rename command ID in all three Revit host implementations. A guarded RVT in Admin OFF now cancels every Project Browser rename command before edit/commit instead of depending on the one-time fallback path.
- Each successful native command binding now retains its command ID, and runtime diagnostics report `projectBrowserRenameBinding=true/false`. This makes a missing version-specific binding visible in the diagnostic log instead of silently failing open.
- Expanded the updater fallback from the duplicated symbol-name parameter set to `SYMBOL_NAME_PARAM`, `ALL_MODEL_TYPE_NAME`, `SYMBOL_FAMILY_NAME_PARAM`, `ALL_MODEL_FAMILY_NAME`, `FAMILY_NAME_PSEUDO_PARAM`, and `ELEM_FAMILY_PARAM` so non-command name changes have broader transaction-level coverage.
- Added static and workflow assertions requiring the real Project Browser command ID, the expanded type-name trigger, and the binding diagnostic in every source host.

### Automated Verification And Deployment

- Revit 2019/2021/2023/2025/2027 Release builds completed without compile errors.
- Static UI/action/native-guard checks passed for all three source hosts: `85` generated actions, `264` exact routes, `65` prefix routes, and `11` browser functions per host.
- Workflow audit passed at `artifacts\family-browser-workflow-audit\20260720-225157`.
- Focused IE `WebBrowser` protected-file scenarios passed in Korean and English for 2019/2023/2025/2027 at `artifacts\family-browser-ui-audit\20260720-225216`; Revit 2021 remains `SKIP runtime-not-installed`.
- Stage and installed add-in verification report `OK` for all five target versions. Stage and ProgramData host DLL hashes match exactly: 2019/2021/2023 `F396BEBFFA0C04C31E667D0443C2FAFE5AA76C6186E7C5939EC5DC68B390036D`; 2025 `29805555C163F3F100D5D861865E670CF4612BF51DE301BDC9240A12D70C1AF3`; 2027 `3CA9A82C4BD9A221B79F9850C6ECB1BD4D963F7A6217DD838EC052E3D351AE5C`.

### Runtime Acceptance Check

- `Needs User Revit Check`: Revit was intentionally not launched automatically because of the reported license error.
- Fully restart Revit so startup command bindings are registered. In a guarded RVT, select Admin OFF and rename the same type and several different Family/System Types at least three times. Every attempt must be blocked before the name commits. Select Admin ON and verify renaming is restored immediately.

## 2026-07-20 First-Attempt Type Rename And Deferred Ribbon Settlement

Status: `Complete / Needs User Revit Check`

### Runtime Evidence And Definitive Root Cause

- `dashboard-runtime-20260720.log` proved Admin OFF propagated immediately inside the guard: `targeted=True`, `canLoad=False`, `canEdit=False`, `canTypes=False`, `updater=True`, and `ribbonState=disabled`. The policy and current Admin state were not the failing layer.
- Revit journal `journal.0158.txt` and the matching `SecurityAudit` entries showed a repeatable per-element pattern: the first rename of a previously unseen System Type could commit, while the second rename of that same element was blocked. Different new element IDs repeated the same first-pass behavior.
- `ProtectedElementIndexes` is intentionally populated incrementally. The Modified branch previously recorded a violation only when a previous index row existed and differed. The first modification of every element missing from the partial index was therefore treated as harmless, then learned by the index, which made only later attempts block.
- Pressing Refresh did not repair the policy. It merely refreshed guard/UI state and changed which element IDs had already been learned. This explains why one type could show an error while the next never-before-seen type still changed once.
- Revit also reapplied its native ribbon state after the modeless Admin transition. The guard disabled `Load Family` immediately, but Revit could visually enable it again until Refresh forced another availability pass.

### Correction

- A Modified Family/Type missing from the guard index now fails closed. `ShouldRecordProtectedChange(...)` blocks an unseen first modification, ignores only a known unchanged element, and blocks a known changed element.
- While Admin Mode is ON and edits are permitted, successful Family/Type changes now update the partial guard index. Returning to OFF therefore compares against the latest authorized state instead of a stale baseline.
- Native `Load Family` state is applied immediately and scheduled for three subsequent Revit `Idling` settlement passes. This reasserts the protected state after Revit finishes its own ribbon refresh, without requiring the browser Refresh button.
- Runtime diagnostics now include `pendingRibbonRefreshPasses` so a delayed ribbon settlement can be distinguished from a policy mismatch.

### Automated Verification And Deployment

- Source checkpoint: `_backups\admin-off-first-attempt-and-ribbon-idle-20260720-2302`.
- Revit 2019/2021/2023/2025/2027 Release builds completed without compile errors.
- Static UI/action/native-guard checks passed for all three source hosts: `85` generated actions, `264` exact routes, `65` prefix routes, and `11` browser functions per host.
- Workflow audit passed at `artifacts\family-browser-workflow-audit\20260720-230921`, including the missing-baseline fail-closed path, Admin ON baseline synchronization, and deferred ribbon settlement.
- Real assembly/IE `WebBrowser` regression passed at `artifacts\family-browser-ui-audit\20260720-230938` for Korean and English protected-file scenarios on Revit 2019/2023/2025/2027. The reflection truth table verifies unseen first Modified = block, known unchanged Modified = ignore, known changed Modified = block, and unseen Added = block. Revit 2021 remains `SKIP runtime-not-installed` on this PC.
- Stage and installed verification report `OK` for all five versions. Stage and ProgramData host DLL SHA-256 match exactly: 2019/2021/2023 `98C25F2F24D89F150C9347ED422ACF78CB7511F4F93628DFACFA7AF110A124D4`; 2025 `14E6F03D74B59897A14BCFA558CEEB2C3C3115CE62E41A7533E45E5E526E8FA0`; 2027 `BB9B6AEA7A31999A2E65F5BBC2AA5F2281798C0F8B7E768201EABCF4D18E756A`.
- Revit was not launched automatically because the user reported license errors from automatic launches.

### Runtime Acceptance Check

- `Needs User Revit Check`: fully restart Revit so the new startup `Idling` subscription and updater assembly are loaded.
- Open the registered File Guard RVT, switch Admin Mode OFF, and confirm native `Load Family` becomes disabled without pressing Refresh.
- Choose at least three different Family/System Types that have not been renamed during this Revit session. The very first rename attempt for every type must be blocked, and repeated attempts must also remain blocked.
- Switch Admin Mode ON and confirm allowed changes work; switch OFF again and confirm the newly authorized state is protected immediately without Refresh.

## 2026-07-20 Admin OFF Live Revit UI Settlement

Status: `Source/Stage/ProgramData Complete / Revit 2019 Runtime PASS`

### Runtime Evidence

- Latest Revit 2019 journal `journal.0159.txt` confirms every tested Pipe, Duct, and Duct System type rename posted the Family Browser protected-change failure and reached `EndOrAbortUndoTransaction`. The database rollback guard is working.
- The matching `SecurityAudit` records at 23:14:06 through 23:14:28 show five consecutive `BlockedBeforeCommit` results. The reported delayed name was therefore the Project Browser inline-label cache, not a committed rename.
- Startup runtime diagnostics restored `requested=False`, `allowed=True`, and `effective=False`, but the preload path called `NotifyAdminModeChanged(..., refreshUiNow: false)` before the latest change. It seeded the correct policy and state without scheduling a native ribbon settlement pass, so Refresh was the next action that visibly disabled `Load Family`.
- First-attempt fail-closed protection could block a type that was absent from the partial index, but no old name was available for the existing pre-commit restoration routine. Revit rolled the transaction back correctly while the Project Browser kept showing the typed label until its next UI command.

### Correction

- `NotifyAdminModeChanged` now schedules native `Load Family` ribbon settlement even when direct UI access is intentionally deferred. Startup and management-folder preload paths remain modeless-safe and apply the saved OFF state on the next Revit `Idling` callback without a Refresh click.
- A protected Admin OFF document now builds one complete, lightweight Family/ElementType name baseline in Revit Idle context. A complete baseline is tracked separately from the previous partial changed-element index and is associated with the current Revit `Document` instance.
- The existing updater can now restore the original Family/Type name inside the blocked transaction before posting the failure, including the first attempted rename in the session.
- Protected rollback processing schedules two post-dialog Revit UI settlement passes and calls `UIDocument.RefreshActiveView()` after the rollback dialog closes. This supplements the pre-commit name restoration without creating an automatic undo/redo transaction.
- Runtime diagnostics now expose `protectedElementBaselineComplete`, `pendingProtectedElementBaselineRefresh`, and `pendingPostRollbackUiRefreshPasses`.

### Verification

- Source checkpoint: `_backups\admin-off-live-ui-settlement-20260720-232219`.
- Static UI/action/native-guard checks passed for all three source hosts: `85` generated actions, `264` exact routes, `65` prefix routes, and `11` browser functions per host.
- Workflow audit passed at `artifacts\family-browser-workflow-audit\20260720-232915`.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage verification completed without compile errors.
- Focused protected-file UI harness passed in Korean and English for 2019/2023/2025/2027 at `artifacts\family-browser-ui-audit\20260720-233015`; Revit 2021 remains `SKIP runtime-not-installed` on this PC.
- Staged host DLL SHA-256: 2019/2021/2023 `DE1462F66A43CCA61FFC80BC9BD056DFEDCEECC70CD4E0DC7947443EB797C15E`; 2025 `18F233693954552F32702F0B32C7BCC8C7790CBD32E66E5EA31FC085BA1BDFED`; 2027 `A575F9C45EABE20400EF389100F1E923734F3D7429285EBA098FB130CE6460FB`.

### Deployment And Runtime Boundary

- The previously pending ProgramData replacement was completed after Revit closed. `Verify-FamilyBrowserRecovered.ps1 -Installed` reports `OK` for Revit 2019/2021/2023/2025/2027, with the installed payload matching the staged revision.
- Revit 2019 was launched from `C:\Program Files\Autodesk\Revit 2019\Revit.exe`, and `C:\Users\kkyki\Desktop\Codex Project\TEST\ProjectKKK.rvt` was opened for the live check. The Family Browser was opened with saved Admin Mode OFF; native `Load Family` became disabled without pressing the browser Refresh button.
- Admin Mode ON immediately enabled native `Load Family`. Switching OFF immediately disabled it again without Refresh, and clicking the disabled ribbon control did not open a file dialog.
- Three consecutive Project Browser type rename attempts were executed: `KKY_BLOCK_TEST_1`, `KKY_BLOCK_TEST_2`, and `KKY_BLOCK_TEST_3`. The first and repeated attempts were all blocked, and the original `Default` label was visible again before the Revit failure dialog was dismissed.
- Revit journal evidence is `%LOCALAPPDATA%\Autodesk\Revit\Autodesk Revit 2019\Journals\journal.0162.txt`; it records all three attempted names and the corresponding transaction abort path.
- Security audit evidence is `D:\TEST\RevitVersions\Revit2019\Projects\387098eea37a35c9cee6c1fb74fcb14f\SecurityAudit\20260720-234546-native-change.log`, `20260720-234726-native-change.log`, and `20260720-234839-native-change.log`. Every record reports `RestoredBeforeCommit | BlockedByNativeGuard`.
- A loadable Annotation Symbols family node was also inspected. Revit 2019 did not expose a Rename command for that family node, so no valid native family-name mutation route existed in this fixture; protected ElementType rename coverage was completed instead.
- The test document was closed without saving. Revit and the Family Browser were both closed at the end of the run.

## 2026-07-21 Browser-Independent Admin OFF Cold Start Guard

Status: `Source/Stage/ProgramData Complete / Revit 2019 Runtime PASS`

### Root Cause

- The previous runtime PASS began after the Family Browser window had opened. The dashboard background startup resolved the homepage/managed-folder path and loaded the shared File Guard policy, so the native guard received the correct saved Admin OFF state at that point.
- Before the first Browser open, `HostWorkspacePathResolver.ResolveRoot()` could still be empty because `FamilyBrowserMachineConfigStore` is runtime-only. A cold Revit session therefore knew that Admin Mode was saved as OFF but did not yet know which shared policy protected `ProjectKKK.rvt`; native `Load Family` could remain enabled until Browser initialization.
- This was a policy-bootstrap ordering gap, not a File Guard match, relative/absolute path, or rename rollback failure.

### Correction

- All three Revit host `App.cs` implementations now queue a one-time native-guard policy preload from the first non-null `ViewActivated` document.
- Revit document title, model path, and central path are captured on the Revit UI thread. Homepage/cached managed-path probing and shared policy file reads then run in a background task.
- The cold preload applies a persisted manual override when present, otherwise calls the path-only deployment bootstrap. It intentionally does not run the full Browser startup, scan the model, build browser rows, or mark the full dashboard bootstrap cache complete.
- The resolved policy and persisted Admin ON/OFF choice are passed to `FamilyBrowserNativeCommandGuardService.NotifyAdminModeChanged(..., refreshUiNow: false)`. Revit ribbon settlement and protected-name baseline work remain scheduled through the existing Revit `Idling` handler, so no Revit API work is performed from the background task.
- The Browser's existing full bootstrap remains the retry/fallback path when the cold managed-path lookup is unavailable.

### Automated Verification And Deployment

- Source checkpoint: `_backups\admin-off-cold-start-policy-preload-20260720-2359`.
- Workflow audit passed at `artifacts\family-browser-workflow-audit\20260721-000626`. New assertions require `HandleViewActivated -> QueueNativeGuardPolicyPreload`, one-time interlocked startup, path-only background bootstrap, and Idling-based Admin state delivery in all three host sources.
- UI static/action/contract checks passed for every source host: `85` generated actions, `264` exact routes, `65` prefix routes, and `11` browser functions.
- Revit 2019/2021/2023/2025/2027 Release build, Stage verification, elevated ProgramData deployment, and installed verification all passed.
- Stage and ProgramData host DLL SHA-256 match exactly: 2019/2021/2023 `6BCD2F7ED1B73AE998975ACCB36B993A27ABB77A3E4CB00D061F3B284FE5C778`; 2025 `0DEDEF57B82E4A0AC92506DACF4B50B929D2F1BA052B316A0A2E5CA4E5CB6173`; 2027 `A4E4780A6C07EE00A4D4DF02F743AA2BD8AE5B46BF79646C5F0A7E8BBF20E155`.

### Revit 2019 Runtime Verification

- Revit 2019 was cold-started with `C:\Users\kkyki\Desktop\Codex Project\TEST\ProjectKKK.rvt` while `%LOCALAPPDATA%\KKY\FamilyBrowser\Settings\admin-mode.txt` was `off`.
- Before the Family Browser window was opened, Revit's Insert-tab `Load Family` control was already disabled. Clicking the disabled control did not open a file picker.
- The Family Browser was then opened and displayed the persisted Admin OFF state. Native `Load Family` remained disabled.
- Two consecutive Project Browser System Type rename attempts were made without Refresh: `Exhaust Air -> KKY_COLD_BLOCK_TEST` and `Return Air -> KKY_COLD_BLOCK_TEST_2`. In both cases the original name was visibly restored before the Revit failure dialog was dismissed.
- Security audit evidence is `D:\TEST\RevitVersions\Revit2019\Projects\387098eea37a35c9cee6c1fb74fcb14f\SecurityAudit\20260721-001544-native-change.log` and `20260721-001645-native-change.log`; both report `RestoredBeforeCommit | BlockedByNativeGuard`.
- Revit journal evidence is `%LOCALAPPDATA%\Autodesk\Revit\Autodesk Revit 2019\Journals\journal.0164.txt`; it records both attempted names and `Cancel, IDABORT` transaction termination.
- The test document was closed without saving. The successful test Revit instance and the earlier licensing-error startup process were both closed; no Revit process remained after verification.

## 2026-07-21 File-Scoped Element Change Tracking And Project History

Status: `Source/Stage/ProgramData Complete / Needs User Revit Check`

### Policy Decision

- Element creation, modification, and deletion tracking is configured per registered RVT in `Permissions / Guard`. This is safer than an unconditional global tracker because test, archive, template, and excluded projects can be opted out explicitly.
- `Track element changes` / `요소 생성·수정·삭제 추적` is checked by default whenever a new RVT is registered or an existing target is re-added.
- Legacy File Guard targets that predate this field migrate to tracking ON. They do not silently become untracked after upgrade.
- When registered File Guard targets exist, only the matching checked RVT is tracked. A matching unchecked RVT is excluded, and unrelated RVTs are not pulled into that policy.
- The administrator tracking switch remains available as a bulk ON/OFF control for all registered RVTs. Per-file choices are made in `Permissions / Guard`.

### Save Boundary And Startup Recovery

- Native guard managed-folder preload now notifies the element tracker and queues the first baseline on the Revit UI thread. Opening the dashboard is no longer required before tracking can start.
- `DocumentSaving` and `DocumentSavingAs` now prepare or recover a missing session before the commit boundary. A late start is recorded as a coverage gap instead of silently dropping the Save.
- An existing session with an unresolved or changed management root is promoted/rebound instead of being returned unchanged.
- Local Save checkpoints, upload-pending records, and immutable synchronized history are merged by stable entry ID in the history view.

### History Browser And Deleted Object Details

- `Current Project History` and `All Project History` are separate actions. The all-project view lists tracked project files and opens the selected project's dedicated history window.
- History can be searched and filtered by Created, Modified, Deleted, and Created-then-deleted. Deleted objects retain the last known category, family, and type when those values were observed before deletion.
- The dedicated table exposes: time, Revit user, change kind, element ID, category, family, type, transaction, PC, commit kind, storage status, local Save time, first/last observed time, Windows user, summary, attribution, policy state, and integrity state.
- Excel is generated only when the user selects `Export Excel`; normal Save, scan, and history viewing do not create diagnostic workbooks automatically.

### Automated Verification

- Source checkpoints: `_backups\element-tracking-save-and-all-project-history-20260721-073026` and `_backups\file-scoped-element-tracking-20260721-075020`.
- UI static/action/contract audit passed for all three hosts: `86` generated actions, `264` exact routes, `65` prefix routes, and `11` browser functions per host.
- Workflow audit passed at `artifacts\family-browser-workflow-audit\20260721-080310`: `905/905` checks across `35` workflows.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage integrity verification passed. Staged host DLL SHA-256: 2019/2021/2023 `5A1BB51D23F376555298630E48EA951A8D52C104FC881E8EE8748C50EDC757C2`; 2025 `BD64F0506A20333EC6418E377E54EFD91ED3C15CF08E3FBB1CB1748D85715B25`; 2027 `79047A870A8A909812F01663DAB7ACF7D3270DDB94CA31C3441DA4AFB8B43C83`.
- Automatic Revit launch was intentionally not performed because the user reported license errors from agent-launched Revit sessions.

### ProgramData Deployment

- The five-version installer was relaunched with the Windows elevated administrator token after the initial filtered-token access denial.
- ProgramData replacement completed for Revit 2019/2021/2023/2025/2027. `Verify-FamilyBrowserRecovered.ps1 -Installed` reports `OK` for every target, including full Stage-to-installed payload hash comparison.
- Installed host DLL SHA-256 matches Stage exactly: 2019/2021/2023 `5A1BB51D23F376555298630E48EA951A8D52C104FC881E8EE8748C50EDC757C2`; 2025 `BD64F0506A20333EC6418E377E54EFD91ED3C15CF08E3FBB1CB1748D85715B25`; 2027 `79047A870A8A909812F01663DAB7ACF7D3270DDB94CA31C3441DA4AFB8B43C83`.

### Runtime Acceptance Check

- `Needs User Revit Check`: restart Revit after installing the staged revision. In `Permissions / Guard`, confirm a newly registered RVT has element tracking checked by default and an old registered RVT also appears checked after migration.
- In a checked standalone RVT, create a wall and Save. `Current Project History` and `All Project History` must show a Created row with time, user, element ID, category, family, and type.
- Modify and Save the element, then delete and Save it. The Deleted filter must show the deleted element ID plus its last known category/family/type immediately after each Save.
- In a workshared RVT, local Save must appear as local/synchronization-pending evidence; successful Synchronize with Central must promote it to confirmed managed history.
- Uncheck tracking for one registered RVT and verify that RVT produces no new tracking session/history while another checked registered RVT continues to do so.

## 2026-07-21 Cold-Start Deletion Baseline Recovery

Status: `Source/Stage/ProgramData Complete / Needs User Revit Check`

### Runtime Evidence And Root Cause

- In `ProjectKKK.rvt`, deleting an existing wall and saving produced Save history but no Deleted row.
- The two records written under `D:\TEST\ElementChangeHistory\C749606C0E7EFE0212BB5270758ED557\20260720` were truthful coverage-gap commits: `CommitKind=Save`, `CoverageGapOnly=true`, and `ChangesCount=0`.
- `D:\TEST\Logs\20260721-081541-963-Element tracking session recovered at Save boundary.log` states that Save began while tracking was enabled but no document baseline session existed.
- The File Guard policy itself was correct: `ProjectKKK.rvt` matched its exact path and both `TrackElementChanges` and the global tracking switch were enabled.
- The defect was startup ordering. The first background managed-path preload was guarded by a permanent one-shot flag. If that first attempt ran before a usable management context or before tracking was enabled, it could not be retried. Dashboard policy load also synchronized the native command guard but did not queue element baseline capture on a valid Revit API event.
- Because the wall was already deleted before the late Save-boundary session was created, its old category, family, type, and element identity cannot be reconstructed retroactively. The existing coverage-gap record is retained as audit evidence.

### Correction

- Managed-path preload now distinguishes `in flight` from `ready`. It is marked ready only after a usable managed data root and policy are resolved, and the in-flight gate is always released in `finally` so an early unresolved attempt can retry.
- Policy resolution now starts from `DocumentOpened` as well as `ViewActivated`, reducing the cold-start window before tracking becomes available.
- Added a shared `RequestDocumentSessionBaselineRefresh()` request channel. Dashboard initial policy load, homepage-path reconnection, project tracking ON, and File Guard save all use this channel instead of scanning the Revit document directly from a modeless form callback.
- All three Revit hosts consume the request on Revit `Idling`, call `BeginDocumentSession(...)`, verify `HasDocumentSession(...)`, and retry a failed baseline up to three times.
- Dashboard diagnostics now emit `element-tracking-policy-sync` with source, enabled state, prior session state, and baseline queue state.

### Automated Verification And Deployment

- Source checkpoint: `_backups\element-tracking-cold-baseline-runtime-fix-20260721-082658`.
- Workflow audit passed at `artifacts\family-browser-workflow-audit\20260721-083112`, including managed-path retry, dashboard-to-Idling baseline routing, File Guard save routing, and real session existence checks.
- Static UI/action/contract checks passed for all three source hosts: `86` generated actions, `265` exact routes, `65` prefix routes, and `11` browser functions per host.
- Revit 2019/2021/2023/2025/2027 Release builds, 2,000-row performance gate, and Stage integrity verification passed.
- The long all-version IE harness exceeded the orchestration time limit after progressing normally. Completed results contained no failures: Revit 2019 `56/56` scenarios and Revit 2023 partial `22/22` scenarios. This change did not modify rendered HTML or action routing.
- Elevated ProgramData deployment and installed verification passed for all five versions. Stage and installed SHA-256 match exactly: 2019/2021/2023 `4C7219F433CFD2E2C4311A6DAB12A7347CFB9E429590FFF2B0FE6DA777E7D623`; 2025 `D2714CA255BED4314061E1B9203D7A921B7DF0E4B230BA5E878ABD9A28D751EC`; 2027 `D58A7654287CDD0A5E373743EC3702A0B435291FC0D624599AB52C906CE5FAA0`.
- Revit was not launched automatically because automatic launches were explicitly prohibited after license errors.

### Runtime Acceptance Check

- `Needs User Revit Check`: fully restart Revit so the newly installed assembly is loaded.
- Open the already checked `ProjectKKK.rvt`. Do not press Refresh and do not require the Family Browser window to be open for the cold-start test.
- Delete one existing wall and Save. `Current Project History` must show a Deleted row with element ID, category, family, and type. The new commit must have `DeletedCount >= 1` and must not be a Save-only `CoverageGapOnly` record.
- Also open Family Browser once and confirm the dashboard runtime log contains `element-tracking-policy-sync ... enabled=True ... baselineQueued=True`.
- If a new `Element tracking session recovered at Save boundary` log appears after this revision, preserve that log and the matching commit JSON for another runtime diagnosis.

## 2026-07-21 Journal-Confirmed Cold-Start Root Race And Standalone Central-Path Guard

Status: `Source/Stage/ProgramData Complete / Needs User Revit Check`

### Journal And Managed-History Evidence

- Revit 2019 journal `%LOCALAPPDATA%\Autodesk\Revit\Autodesk Revit 2019\Journals\journal.0169.txt` records the exact order: `Delete the selection` at line 1338, successful `Delete Selection` at lines 1360-1361 around 08:43:37, Save at line 1364 and 08:43:39, then Family Browser open at line 1429 around 08:43:41.
- The browser open was followed by `This Document is not a workshared document` and `ADocument::getWorksharingCentralModelPath` at lines 1438-1442, including dump `journal.0169.0001.dmp`.
- The central-path exception happened after the deletion and Save. It was a real independent add-in defect and explains the popup/delay, but it did not cause the already-missed deletion record.
- `D:\TEST\Logs\20260721-084338-826-Element tracking session recovered at Save boundary.log` reports `Workspace:` blank and `Save began while tracking was enabled but no document baseline session existed`.
- The matching 08:43:39 and 08:44:25 managed-history files are truthful gap records: `CommitKind=Save`, `CoverageGapOnly=true`, `Changes=0`, and `DeletedCount=0`. The deleted wall cannot be reconstructed retroactively.
- Cold-start code tracing found the direct cause: the active machine configuration and homepage bootstrap cache were runtime-only. `DocumentOpened` could return an editable document before the homepage-managed root and shared tracking policy were resolved. The previous Idling retry fix could establish a later baseline, but could not observe a deletion that happened during this startup window.

### Correction

- Backup: `_backups\journal-baseline-centralpath-fix-20260721-090245`.
- All three `FamilyBrowserMachineConfigStore.cs` hosts now retain the last successfully accepted managed-policy path in `%LOCALAPPDATA%\KKY\FamilyBrowser\Settings\last-known-managed-policy-path.txt`. The hint is written only after the normal managed-path validation/setter succeeds; homepage refresh remains authoritative and can replace it.
- All three host `App.cs` implementations now restore the persisted TEST override or last verified managed path inside `DocumentOpened`, apply the shared tracking policy, and only then call `BeginDocumentSession(...)`.
- On a first run with no usable path hint, `DocumentOpened` performs the homepage path-only bootstrap before returning the project for editing. A later background refresh still checks for a changed homepage path.
- `FamilyBrowserDashboardHtmlForm.TryGetCentralPath(...)` and `StandardRvtChangeCandidateService.BuildDocumentComparePaths(...)` now call `GetWorksharingCentralModelPath()` only when `Document.IsWorkshared` is true. All remaining central-path call sites were statically reviewed and already had the same guard.
- The UI harness route check now distinguishes exact `project-element-change-history` and `project-element-change-history-all` actions. Its former prefix comparison falsely reported the intentional current/all-project buttons as duplicates.

### Automated Verification

- Workflow audit passed `958/958` checks across 35 workflows at `artifacts\family-browser-workflow-audit\20260721-090740`.
- Integrated no-harness quality gate passed static/action/contract, nested-family propagation, authoritative System Type apply, workflow tracking persistence, five-target Stage verification, and 2,000-row performance/cache at `artifacts\family-browser-ui-audit\20260721-092200-journal-fix-quality-gate`.
- Revit 2019/2021/2023/2025/2027 Release builds completed without compile errors. Stage DLL SHA-256: 2019/2021/2023 `58AB6880C6B53B54251A9F3F6918F152C42222DE3816B68FB2535E8EE275194B`; 2025 `DC0F54C74BF2E7C56EB8037430C7BF15846BB794F27941F9C76709BD776CADB9`; 2027 `8C64ABB415DE233C5E0A64425FC416A0B7EACEBBAB52027C93DBD054502EC9A9`.
- Targeted IE `WebBrowser` regression passed all 8 Korean/English message, administrator-layout, and 1920px viewport scenarios at `artifacts\family-browser-ui-audit\20260721-090834-quality-gate\harness-targeted-2019` after correcting the exact-route harness assertion. A post-change static/action/contract rerun also passed with `86` generated actions, `265` exact routes, `65` prefix routes, and `11` browser functions per host.
- The all-version UI harness exceeded the 10-minute orchestration limit after completing Revit 2019 and 2023 plus part of 2025. The completed 2019/2023 failures were only the harness prefix false positive described above; the focused rerun passed. No full-harness PASS is claimed for this revision.

### Deployment And Runtime Boundary

- The initial direct write probe correctly showed that the Addins roots permit new files, but the existing manifests and payload files are owned by `BUILTIN\Administrators` and grant `BUILTIN\Users` read-only access. Direct overwrite/delete therefore failed under the medium-integrity Codex shell even though root-level file creation succeeded.
- ProgramData checkpoint before replacement: `_backups\programdata-before-journal-root-fix-20260721-0932` (`17` files).
- The verified installer was then run through Windows elevation and completed with exit code `0`. `Verify-FamilyBrowserRecovered.ps1 -Installed` reports `OK` for Revit 2019/2021/2023/2025/2027, including exact Stage-to-installed payload comparison.
- Installed host DLL SHA-256 now matches Stage exactly: 2019/2021/2023 `58AB6880C6B53B54251A9F3F6918F152C42222DE3816B68FB2535E8EE275194B`; 2025 `DC0F54C74BF2E7C56EB8037430C7BF15846BB794F27941F9C76709BD776CADB9`; 2027 `8C64ABB415DE233C5E0A64425FC416A0B7EACEBBAB52027C93DBD054502EC9A9`.
- `Needs User Revit Check`: fully restart Revit, open `ProjectKKK.rvt`, do not open or refresh Family Browser, delete an existing wall, and Save. The new history entry must have `DeletedCount >= 1`, must not be `CoverageGapOnly`, and no new `Element tracking session recovered at Save boundary` log should appear.
- Opening Family Browser on the standalone `ProjectKKK.rvt` must no longer add a `getWorksharingCentralModelPath` warning or dump to the journal.

## 2026-07-21 Journal 0171 Cross-Callback Document Identity Split

Status: `Source/Stage/ProgramData Complete / Needs User Revit Check`

### Runtime Evidence And Root Cause

- Revit 2019 journal `%LOCALAPPDATA%\Autodesk\Revit\Autodesk Revit 2019\Journals\journal.0171.txt` proves the add-in registered `DocumentOpened`, `DocumentChanged`, `DocumentSaving`, and `DocumentSaved` successfully at lines 660-665.
- The first test created a wall at line 1293, raised document change processing around line 1358, deleted the selected wall at lines 1469/1488, and saved at line 1502. The second test repeated create/change at lines 1627/1677/1692, delete at lines 1817/1836, and save at line 1852.
- The two matching managed-history files, `D:\TEST\ElementChangeHistory\C749606C0E7EFE0212BB5270758ED557\20260721\0242fa41297347529910f0ed9f4f88b4.json` and `2723813e994448e68928a72fe5fc5ebe.json`, were both `CommitKind=Save`, `CoverageGapOnly=true`, `ActivityCount=0`, and had zero Created/Modified/Deleted rows.
- `D:\TEST\Logs\20260721-095714-412-Element tracking session recovered at Save boundary.log` confirms that Save saw no baseline session. Dashboard diagnostics also remained `sessionReady=False` / `BaselineMissing` even after policy load and after the Browser was open.
- The policy was not the failure: `D:\TEST\Config\standard-policy.json` had both the global element tracker and the exact `ProjectKKK.rvt` File Guard target enabled.
- The defect was `BuildRuntimeKey(...)`. It used a `ConditionalWeakTable<Document,...>` value initialized with a random GUID. Revit can provide different managed `Document` wrapper objects for `DocumentOpened`, `DocumentChanged`, and `DocumentSaving`; the same RVT therefore split into unrelated tracking sessions. Save recovered a new empty session instead of committing the session that received the create/delete events.

### Correction

- Source checkpoint: `_backups\element-tracking-stable-document-key-20260721-100744`.
- `FamilyBrowserElementChangeTrackingService.BuildRuntimeKey(...)` now creates a callback-stable identity from the normalized local RVT path first, then the normalized project/central identity, then a deterministic unsaved-document identity.
- The per-wrapper weak-table cache remains so the same wrapper keeps its identity through a Save As boundary, but different wrappers for the same saved RVT now derive the same key instead of unrelated GUIDs.
- A random identity is retained only as the final fallback when Revit exposes no usable path, project identity, or title.
- Workflow assertions now require stable local/project identity derivation and reject the former parameterless random-identity construction.

### Automated Verification And Deployment

- Workflow audit passed `962/962` checks across `35` workflows at `artifacts\family-browser-workflow-audit\20260721-100821`.
- Static UI/action/contract checks passed for all three source hosts: `86` generated actions, `265` exact routes, `65` prefix routes, and `11` browser functions per host.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage verification passed. Existing unrelated compiler warnings remain; this change introduced no compile error.
- ProgramData checkpoint: `_backups\programdata-before-element-tracking-stable-key-20260721-101042` (`17` files).
- Elevated ProgramData deployment and installed verification passed for all five versions.
- Stage and installed host DLL SHA-256 match exactly: 2019/2021/2023 `8FE346CBF611823F981E3A3A8CF5C1E06764D95DF73FA47E27CCC7082ED18804`; 2025 `626E706FB005CBBBAE920E08E52D9DE881C379FB094DA77BC1E55522A4F32E24`; 2027 `0FAE366C1D566BD1A3BA046BE4C17BCB9FC64F12821C41D530AFA4BA72115891`.
- Revit was not launched automatically, as explicitly requested after the license failures.

### Runtime Acceptance Check

- `Needs User Revit Check`: fully restart Revit so the newly installed DLL is loaded, then open `ProjectKKK.rvt`. Browser open/Refresh must not be required.
- Create one wall and Save before deleting it. The new history commit must contain `CreatedCount >= 1` and a Created row.
- Delete one existing wall and Save. The next history commit must contain `DeletedCount >= 1` and a Deleted row with element ID plus last-known category/family/type.
- Neither commit may be `CoverageGapOnly`, and no new `Element tracking session recovered at Save boundary` diagnostic should be created.

### Installer Artifact

- Version `1.0` five-target installer created from this revision: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_element-tracking-fix-20260721-1017_Setup.exe`.
- Included Stage targets: Revit 2019/2021/2023/2025/2027; payload manifest contains `17` files and passed Stage verification.
- Installer size: `3,803,183` bytes (`3.63 MB`). SHA-256: `C73823474748CA91E896E82C246862F3536EF6F3DD2EA19BEC382E67E65A2829`.
- The generated `.sha256.txt` value matches the installer binary. Existing installer artifacts were preserved, and no mail-sized package was generated for this installer-only request.

## 2026-07-21 Secure Guided Product Auto Update

Status: `Source/Stage/ProgramData/Installer Complete / Needs User Revit Check`

### Implemented Update Flow

- `Update Check` still reads `https://update.zerokky.com/Release/family-browser/latest.json`, but a newer version now offers `Download Update` instead of only opening the homepage.
- The installer is downloaded on a worker thread to `%LOCALAPPDATA%\KKY\FamilyBrowser\Updates\<version>` so the Revit UI remains responsive.
- Privileged execution is allowed only for an absolute HTTPS `.exe` URL on the exact `update.zerokky.com` feed host. HTTP, another host, a non-default HTTPS port, user-info, fragments, and non-EXE targets are rejected.
- Redirect destination, 128 KB minimum/256 MB maximum size, `MZ` header, and the manifest SHA-256 are validated before caching. A cached installer is reused only after the same validation passes again.
- The downloaded file is revalidated immediately before launch. The verified installer is started with Windows `runas`; the user must approve UAC, save or synchronize open projects, close Revit, complete installation, and start Revit again.
- The installer explicitly uses `CloseApplications=yes` and `RestartApplications=no`. The updater never force-kills or automatically restarts Revit.
- Update failures do not execute the file and offer the homepage as a manual recovery route.

### Release And Version Guards

- The current product version remains exactly `1.0`. No `1.0.1` or `1.1` package was published or activated.
- `Build-FamilyBrowserInstaller.ps1` now reads `CurrentProductVersion` from the runtime source and blocks the build when it differs from `-Version`. A deliberate `-Version 1.1` test failed closed with `source=1.0 package=1.1`.
- Future `1.1` release order is fixed: change the source version to `1.1`; build with `-Version 1.1`; upload the final installer; verify its bytes and SHA-256; update `latest.json` last with version, HTTPS URL, SHA-256, notes, and publish time.
- Users who install this updater-enabled 1.0 bridge can then check for 1.1, download and verify it, and start the installer from Family Browser.
- Important compatibility boundary: already distributed older 1.0 binaries do not contain this downloader. Because their version is also `1.0`, they cannot receive the bridge as a same-version automatic update. They need one manual install of this bridge build; without it, a later 1.1 check retains the older homepage/manual behavior.

### Automated And Live Verification

- Source checkpoint: `_backups\family-browser-secure-auto-update-20260721-102707`.
- Static UI/action/contract audit passed: `86` generated actions, `265` exact routes, `65` prefix routes, and `11` browser functions for every host source.
- Offline updater primitives passed in the real 2019/2023/2025/2027 host assemblies: trusted HTTPS URL accepted; HTTP/foreign-host/non-EXE URLs rejected; intact MZ+SHA accepted; wrong SHA and non-MZ bytes rejected. Targeted IE harness report: `artifacts\family-browser-ui-audit\20260721-secure-auto-update-targeted`; `24/24` executed scenarios passed, Revit 2021 runtime was `SKIP runtime-not-installed`.
- Workflow audit passed `962/962` checks across `35` workflows at `artifacts\family-browser-workflow-audit\20260721-103718`.
- Live server audit passed and is recorded at `artifacts\family-browser-update-audit\summary.json`: feed version `1.0`, installer `3,780,528` bytes, MZ valid, SHA-256 `42A21F2410BD1A709B22D0D8E44FC33F57C9BB5CE1786D2C4CB57E773B4C5AF7` matching the feed.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage verification passed. Stage and installed host DLL hashes match: 2019/2021/2023 `E5908F058A071247304201DF4C88B3B35D38DB55D6F2DCDB65438231714F78BF`; 2025 `BBAF43DE96D30F95EA6367F4DC7507CF516966D9B4723E3F3F4EE41EC920B069`; 2027 `3F5EA7B1D23F2CE54AC1458AEF8F6CE964D786C91E2F59A89EE61A90BB26A50D`.

### Deployment And Artifact

- ProgramData checkpoint: `_backups\programdata-before-secure-auto-update-20260721-104001` (`17` files).
- Elevated ProgramData deployment and installed payload verification passed for Revit 2019/2021/2023/2025/2027. Revit was not launched automatically.
- Updater-enabled bridge installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_secure-auto-update-20260721-1041_Setup.exe`.
- Installer size: `3,808,802` bytes. MZ and metadata hash verification passed. SHA-256: `6461949D1D707B1EABF7BABE9089F04F04E65A0C0A63489BEC1C345F2E14EF6D`.

### Security And Runtime Boundary

- HTTPS same-host enforcement plus SHA-256 detects transport corruption, wrong files, stale cache, and tampering after download. The hash is supplied by the same HTTPS server, so this does not protect against full compromise of the update server itself. Authenticode signing or a separately pinned signing key remains a future hardening option.
- `Needs User Revit Check`: after installing the bridge and restarting Revit, `Update Check` must report current `1.0` while the live feed remains at `1.0`.
- A true end-to-end newer-version test requires a staged feed or real 1.1 artifact. It must verify: update prompt, background download, hash pass, UAC launch, locked-Revit guidance, no automatic Revit restart, and successful 1.1 startup after installation.

## 2026-07-21 Dedicated History Navigation And Per-RVT Tracking Scope

Status: `Source/Stage/ProgramData Complete / Needs User Revit Check`

### Navigation And Policy Correction

- Source checkpoint: `_backups\history-navigation-file-policy-20260721-104813` (`15` files).
- The administrator sidebar now has a dedicated `History / 이력 관리` group with two direct actions: `Current Project History / 현재 프로젝트 이력` and `All Project History / 전체 프로젝트 이력`.
- The duplicate global `Track project element creation, modification, and deletion` checkbox and its `project-element-change-tracking/*` route were removed from Standards settings.
- Element tracking now has one authoritative configuration surface: an RVT must be an enabled target in `Permissions / Guard`, and that target's `Element Change Tracking / 요소 변경 추적` checkbox must be enabled.
- The legacy `TrackProjectElementChanges` JSON field remains only as a compatibility mirror written by File Guard saves. It can no longer enable an unregistered RVT or bypass a target checkbox.
- An active project with no stored history is inserted into the all-project picker only when it currently matches a registered, checked File Guard target. Existing immutable history remains viewable after a file is later removed from the policy.

### History Window Guidance

- Both the all-project selection window and the current/selected-project history table now show a visible scope notice: only RVTs registered in Permissions / Guard with Element Change Tracking enabled receive new history.
- The empty-history row now directs the administrator to register the RVT and enable its per-file tracking checkbox instead of referring to the removed global option.
- Historical records are not deleted or hidden when a policy target is removed; the notice explicitly distinguishes new recording scope from retained audit history.

### Automated Verification

- Static UI/action/contract checks passed for all three source hosts: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host. The one-action/one-prefix reduction is the intentional removal of the global tracking toggle.
- Workflow audit passed `958/958` checks across `35` workflows at `artifacts\family-browser-workflow-audit\20260721-105637`.
- The UI harness now verifies three policy cases against the compiled host assemblies: matching registered target checked = enabled; matching target unchecked = disabled; unregistered RVT = disabled.
- Targeted IE `WebBrowser` checks passed with zero failures for Korean/English administrator Home, Standards settings, and 1920px administrator scenarios on Revit 2019/2023/2025/2027 at `artifacts\family-browser-ui-audit\20260721-105739-history-navigation`.
- Revit 2021 Stage/package verification passed; runtime UI smoke is `SKIP runtime-not-installed` because Revit 2021 is not installed on this PC.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage verification passed. Existing unrelated compiler warnings remain; this change introduced no compile error.

### ProgramData Deployment

- ProgramData checkpoint: `_backups\programdata-before-history-navigation-20260721-110027` (`17` files).
- Elevated ProgramData deployment completed for Revit 2019/2021/2023/2025/2027, and installed-payload verification passed for every target.
- Stage and installed host DLL SHA-256 match exactly: 2019/2021/2023 `8E7199FD037283AA54FB5086688B1D0FA9C569AAF585FA72C3C45B3E2BF73447`; 2025 `A7FE0C7ABBABCC65408ACEDE9A7F3520A47B21CE38D094AF6F6D2F6A66B24D96`; 2027 `3BFBF74497A714259D6E9EA73F57268AEE0FE95971BCE733C6B297AC385A3FBF`.
- Revit was not launched automatically, following the explicit license-error constraint.

### Runtime Acceptance Check

- `Needs User Revit Check`: restart Revit, open Family Browser as an administrator, and confirm the left sidebar shows the new History group while Standards settings no longer shows a global element-tracking checkbox.
- Register one RVT in Permissions / Guard with Element Change Tracking checked, edit and Save or Synchronize it, then confirm the entry appears through both history menus.
- Clear that RVT's tracking checkbox and confirm subsequent edits create no new history. An unrelated, unregistered RVT must also create no new history.
- Previously committed history must remain available in All Project History after the RVT is removed from Permissions / Guard.

## 2026-07-21 Shared Parameter And Grid Detail Tracking

Status: `Source/Stage/ProgramData Complete / Needs User Revit Check`

### Implemented Tracking Scope

- Source checkpoint: `_backups\shared-parameter-grid-tracking-20260721-112453` (`8` files).
- Project and shared parameter definitions are now explicit tracked elements even though Revit exposes them without an ordinary model category.
- A parameter history state contains its name, shared-parameter GUID, instance/type/unbound binding, bound category names and IDs, parameter group, data type, and varies-across-groups state.
- Save and Synchronize boundaries re-read the project parameter definitions and full `BindingMap`. This recovers final binding changes that Revit may not identify with a reliable parameter element ID in `DocumentChanged`.
- Grid history now contains the grid name and type, model curve signature in millimetres, model extents, pinned state, and workset. Created, modified, deleted, and created-then-deleted records retain the last trustworthy metadata available to the observing client.
- Current Project History, All Project History, and their explicit Excel export now have a separate Name column and readable Korean/English summaries for parameter binding/category changes and grid curve/extent/pin changes.

### History Schema And Compatibility

- New element-history records use schema `6` and integrity version `5`. The protected checksum now includes the item name/tracking kind plus every added parameter and grid field.
- Integrity versions `1` through `4` remain frozen projections. Existing schema `1` through `5` history and legacy pending envelopes remain readable and are not rewritten as new evidence.
- Automated tamper tests proved that changing a shared-parameter bound category or a grid curve in the stored JSON invalidates the record instead of accepting modified evidence.
- A parameter binding-map read failure at Save or Synchronize creates an explicit commit-boundary coverage gap and diagnostic instead of silently treating parameter metadata as complete.

### Automated Verification

- Workflow audit passed `970/970` checks across `35` workflows at `artifacts\family-browser-workflow-audit\20260721-shared-parameter-grid-final`.
- Static UI/action/contract checks passed for all three host sources: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- Targeted Korean/English IE `WebBrowser` rendering passed `32` executed scenarios with zero failures at `artifacts\family-browser-ui-audit\20260721-shared-parameter-grid`; Revit 2021 UI runtime was `SKIP runtime-not-installed`.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage verification passed. Existing unrelated compiler warnings remain; this work introduced no compile error.

### ProgramData Deployment

- ProgramData checkpoint: `_backups\programdata-before-shared-parameter-grid-20260721-114453` (`17` files, `21,713,738` bytes).
- Elevated ProgramData deployment and installed-payload verification passed for Revit 2019/2021/2023/2025/2027.
- Stage and installed host DLL SHA-256 match exactly: 2019/2021/2023 `74E26DB51A2356B75431FDB7BB618ED4EAF7DE12E810CB2CA92AC10A70972841`; 2025 `6EBAF3388F201896E00672E46AA6E342B58271A1F9E6D443AFB56DE3463B5C2B`; 2027 `30AD71A013C424A0CF56301F6971D5B6F8984731C165B9E872381FA06B982BCB`.
- Revit was not launched automatically because the user will perform runtime testing and automated Revit startup can trigger a licence error on this PC.

### Runtime And Evidence Boundary

- `Needs User Revit Check`: register the test RVT in Permissions / Guard and enable its per-file Element Change Tracking checkbox before testing.
- Add a shared parameter as an instance binding to two categories, then Save or Synchronize. History must show the name, GUID, binding, and both categories. Change the categories or binding kind and verify the row shows old and new values; then delete the definition and verify a Deleted row.
- Create, rename, move, pin/unpin, and delete a grid with a Save between each operation. History must show its element ID, name, curve/extents/pin differences, user, PC, and commit time.
- Workshared local Save protects a pending checkpoint; the record becomes centrally committed only after successful Synchronize with Central. Standalone files commit on successful Save.
- Exact attribution still requires this add-in on every editing workstation. A final-state Save comparison can recover parameter binding differences, but it cannot reconstruct the exact author or exact intermediate sequence from a workstation where the add-in was absent.
- A parameter added and deleted entirely between two boundaries can only be retained when Revit emitted usable `DocumentChanged` evidence. Per-view grid bubble visibility and per-view 2D datum extents are not part of this model-level grid signature.

## 2026-07-21 Element Tracking Correctness Review

Status: `Review Complete / 2 Findings / Needs Fix / Needs User Revit Check`

### Confirmed Tracking Flow

- Revit 2019/2021/2023/2025/2027 hosts subscribe to document open/change/save/save-as/synchronize/reload-latest/close boundaries and route them into the shared element tracking service.
- A pre-change baseline is captured for registered RVTs whose Permissions / Guard target has Element Change Tracking enabled. Final commit classification compares the baseline, observed `DocumentChanged` IDs, and the final state at a successful boundary.
- Standalone RVTs publish immutable history after a successful Save. Workshared local Save writes a protected local checkpoint; a later successful Synchronize with Central publishes that checkpoint to managed immutable history.
- Project/shared parameter definitions and their `BindingMap` metadata are re-read at Save/Synchronize, so binding-kind and bound-category changes can be recovered even when `DocumentChanged` does not identify the binding change reliably.
- Managed history, local upload spool, and local-save checkpoints use destination binding, schema validation, integrity hashes, atomic promotion, revision tokens, and idempotent entry IDs. Invalid or destination-mismatched records are not accepted as confirmed history.

### Actual Evidence On This PC

- The active managed policy is `D:\TEST\Config\standard-policy.json`; its enabled target is `ProjectKKK.rvt` with element tracking enabled.
- Existing `D:\TEST\ElementChangeHistory` evidence includes a successful 2026-07-21 10:17 Save commit for ProjectKKK with two created walls, one modified wall, and one deleted wall. Each row retains element ID, class/category, family, type, commit time, Revit user, Windows user, and machine information.
- The local pending tracking area contains no unresolved element commit or session checkpoint record beyond the zero-byte lock file. The observed ProjectKKK commit was therefore persisted rather than left only in the local queue.
- These existing records use schema 5/integrity 4 because they were produced before the current schema 6/integrity 5 deployment. They prove the generic wall create/modify/delete flow, but do not constitute runtime proof of the newly added parameter/grid fields.

### Finding 1 - Ambiguous Same-Name RVT Scope

- Severity: `High` for a file-scoped audit/guard policy.
- `FamilyBrowserSecurityPolicyService.FindMatchingFileGuardTarget(...)` first compares paths, but if no exact normalized path matches it falls back to the detached-normalized file base name and returns the first enabled target with that name.
- An unrelated `C:\Other\Project.rvt` can therefore inherit the tracking and guard policy registered for `C:\Managed\Project.rvt`. If two registered targets share a filename but have different options, a mapped-drive/UNC/hostname mismatch can also select the wrong target.
- The compiled harness checks an exact/alias target and an unrelated RVT named `OtherModel`; it does not cover an unrelated RVT with the same filename. Required correction: compare canonical/stable path identities first, allow a detached/local lineage fallback only when it is explicit and unambiguous, and report an ambiguous match instead of selecting the first target.

### Finding 2 - Grid Extents Return-Type Mismatch

- Severity: `Medium` because create/delete and curve/pin tracking still work, but the promised extents field is incomplete.
- `SafeGridExtentsSignature(...)` invokes `Grid.GetExtents()` and casts the result to `BoundingBoxXYZ`. Reflection against installed Revit 2019, 2023, 2025, and 2027 APIs confirms that the return type is `Autodesk.Revit.DB.Outline` in every target version.
- The cast therefore yields `null`, so `GridExtentsSignature` remains empty. Required correction: read `Outline.MinimumPoint` and `Outline.MaximumPoint` (with a version-safe fallback) and add a regression check for a non-empty extents signature.

### Automated Review Results

- Workflow audit passed `970/970` checks across `35` workflows at `artifacts\family-browser-workflow-audit\20260721-tracking-review`.
- Static UI/action/contract checks passed for all three source hosts: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- Targeted compiled IE `WebBrowser` checks passed for Korean/English administrator Home and structured result dialogs on Revit 2019/2025/2027 at `artifacts\family-browser-ui-audit\20260721-tracking-review` with zero failures.
- Revit was not launched automatically because runtime launch on this PC can trigger the reported licence error.

### Runtime Boundary

- `Needs User Revit Check`: the current schema 6 build still needs one real Revit pass for shared parameter add/binding change/delete and grid create/rename/move/pin/delete after the two findings are corrected.
- Exact authorship is available only for changes observed by an installed/running add-in. Changes made on another workstation without the add-in can be detected later only as an external final-state difference; the exact author and intermediate transaction sequence cannot be reconstructed.
- Shared-parameter tracking here means definitions and bindings inside an RVT project. It does not monitor edits to the external shared-parameter TXT file or parameter edits made inside an RFA family document.

## 2026-07-21 Browse Trade Column Synchronization

Status: `Source/Stage/ProgramData Complete / Needs User Revit Check`

### Root Cause And Correction

- Source checkpoint: `_backups\20260721-browse-trade-column-sync`.
- The Family/System list data switched to the selected trade correctly, but `DisplayDisciplineLabel(...)` discarded policy slot keys such as `discipline-architecture` before resolving their display names.
- When the built-in discipline parser could not parse that slot key, `ResolvePolicySlot(...)` silently returned the policy-active trade. A browse-only switch from Mechanical/Piping to Architecture could therefore leave the row `Trade / 공종` column labelled with the previous policy-active trade.
- All three hosts now resolve the exact row slot key, discipline name, or custom display name first with `ResolvePolicySlotExact(...)`. The row-label path no longer falls back to an unrelated policy-active trade.
- The correction applies to both Family and System Type row payloads because both use the same discipline-label renderer.

### Regression Coverage

- The audit fixture now reproduces the real mismatch: browse target `discipline-architecture`, policy-active target `Mechanical`, and row keys stored in slot-key form.
- The IE `WebBrowser` harness now compares the selected trade chip label with both the virtual-row `data-discipline` payload and the rendered table's `Trade / 공종` cell.
- Static UI/action/contract checks passed for all three source hosts: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- Targeted Korean/English Family-list checks passed with zero failures for Revit 2019/2023/2025/2027 at `artifacts\family-browser-ui-audit\20260721-142110`. Revit 2021 runtime is `SKIP runtime-not-installed`; its shared 2019/2021/2023 assembly and staged package were verified.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage integrity verification passed.

### ProgramData Deployment

- ProgramData checkpoint: `_backups\programdata-before-browse-trade-column-sync-20260721` (`17` files, `21,891,214` bytes).
- Elevated ProgramData deployment and installed-payload verification passed for Revit 2019/2021/2023/2025/2027.
- Stage and installed host DLL SHA-256 match exactly: 2019/2021/2023 `4B0A54FC0F884224AB501085964A5555FB5641879B10AA10CA4EC6BEAC34289D`; 2025 `80E7533CA0DD608F45E0256D9CA1B658B9BE9EDFE24FE6C0AACF350B6AB30A85`; 2027 `08505B5629ECBDC7147C6A1C2796AF69ABB287AD913A582A7ACB8590DD651625`.
- Revit was not launched automatically because runtime launch on this PC can trigger the reported licence error.

### Runtime Acceptance Check

- `Needs User Revit Check`: register and scan at least two trades, open the Family list, and switch the trade chips below the search box in both directions.
- Confirm that the list content, left trade/category tree, and every visible `공종` column cell change to the same selected trade immediately without pressing Refresh.

## 2026-07-21 File Guard Canonical Path Matching

Status: `Source/Stage/ProgramData Complete / Needs User Network Revit Check`

### Confirmed Root Cause

- Source checkpoint: `_backups\file-guard-canonical-path-match-20260721`; ProgramData checkpoint: `_backups\programdata-before-file-guard-canonical-path-20260721-151631` (`17` files).
- File Guard target selection had two different implementations. `FamilyBrowserNativeCommandGuardService` expanded mapped drives and compared hostname/IP UNC aliases, while `FamilyBrowserSecurityPolicyService` used only `Path.GetFullPath(...)`. Browser Family Load permission, element-history scope, and native Revit command blocking could therefore disagree for the same RVT.
- On the initial Browser render, central-path lookup can be deferred for startup performance. The localized status text `Not checked on startup / 시작 시 확인 안 함` was carried in `CentralPath` and could be mistaken for an already resolved path until Refresh.
- After a path miss, both implementations returned the first enabled target whose detached-normalized filename matched. An unrelated standalone same-name RVT or two same-name policy targets could therefore receive an arbitrary policy.

### Correction

- Added shared `FamilyBrowserFileGuardPathMatcher` and routed both Browser/security decisions and native Revit command decisions through it. Family Load, Family Edit, Family/Type rename, type add/delete, nested-only placement, and per-RVT element tracking now use one target decision.
- Matching order is normalized path, mapped-drive expansion through cached `WNetGetConnection`, exact `share + subfolder + RVT filename` UNC identity for hostname/IP aliases, then Windows file identity for remaining reachable aliases.
- Localized startup status text is rejected as a path. Before central-path refresh, a workshared local/detached document may use filename lineage only when exactly one enabled target matches. Standalone same-name RVTs never inherit that fallback, and multiple same-name targets return an explicit ambiguous no-target result instead of selecting the first row.
- Runtime Guard diagnostics now include model path, resolved central path, match kind, ambiguity/candidate count, and matched policy target so a remote-PC mismatch can be diagnosed from Debug Log evidence.

### Automated Verification

- Static UI/action/contract checks passed for all three source hosts: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- The compiled-host audit now verifies Admin OFF before manual Refresh, hostname/IP UNC equivalence in both Browser and native guards, different-share isolation, unrelated standalone same-name isolation, and ambiguous two-target rejection.
- Korean/English `admin-profile-off-protected-family` IE `WebBrowser` scenarios passed with zero failures for Revit 2019, 2025, and 2027 at `artifacts\family-browser-ui-audit\20260721-file-guard-path-2019`, `...-2025`, and `...-2027`. Revit 2021/2023 use the verified 2019/2021/2023 shared assembly.
- Workflow audit passed `970/970` checks across `35` workflows at `artifacts\family-browser-workflow-audit\20260721-file-guard-canonical-path`.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage verification passed. Existing unrelated compiler warnings remain; this change introduced no compile error.

### ProgramData Deployment And Runtime Boundary

- Elevated ProgramData deployment and installed-payload verification passed for Revit 2019/2021/2023/2025/2027.
- Stage and installed host DLL SHA-256 match exactly: 2019/2021/2023 `DA699A72BBEAA292961646421AB05BBFA6AC67CF148360CBC41F63667F756272`; 2025 `97F1A1BF630C7671FB25C46E98FEAE938552D99850597B06BE7463597F3C59F4`; 2027 `29BC410E11ADD8F9ECE02C6F0798868EB9B71CCEE4E0AAE245B3F6247489AB2F`.
- Revit was not launched automatically because the user will perform runtime testing and automatic startup can trigger the reported licence error.
- `Needs User Network Revit Check`: restart Revit on the affected PC, register the target once through its mapped-drive form, then open the same workshared RVT through its local/UNC form. With an administrator profile and Admin Mode OFF, Family Load plus Family/Type rename and type add/delete must be blocked immediately without pressing Refresh.
- Repeat once with the target registered by hostname UNC and the central reported by IP UNC. Debug Log runtime evidence must show `targeted=True`, `canLoad=False`, `canEdit=False`, `canTypes=False`, and `fileGuardMatch=kind=PathIdentity` (or the explicit unique workshared-name fallback only before the central path is available).

## 2026-07-21 File Guard Canonical Path Installer

Status: `Installer Complete / Needs User Network Revit Check`

- Product version remains exactly `1.0`.
- Rebuilt and Stage-verified Revit 2019/2021/2023/2025/2027 payloads (`17` staged files) before packaging.
- Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0_file-guard-canonical-path-20260721-1521_Setup.exe`
- Installer size: `3,822,569` bytes (`3.65 MB`).
- Installer SHA-256: `989D8C9820ACB4E446257ED5F7637C312960DD4019CDAC5C8105CDE0934727D2`; the independently calculated hash, `latest-build.json`, and companion `.sha256.txt` all match.
- Existing installer artifacts were preserved. A mail-sized package was not generated for this installer-only request.
- Revit was not launched automatically; the affected mapped-drive/UNC/IP File Guard behavior remains `Needs User Network Revit Check`.

## 2026-07-21 Permission / Guard Long-Path Layout And Version 1.0.1

Status: `Source/Stage/ProgramData/Installer Complete / Needs User Revit Check`

### Problem And Correction

- Source checkpoint: `_backups\version-1.0.1-permission-layout-20260721-152901` (`20` files, `6,856,190` bytes).
- The Current Model, File-specific Guard, and Current Model Tracking diagnostics used a three-column card row. Long local, central, and managed-root paths therefore consumed the narrow cards and could distort the Permissions / Guard page.
- The three diagnostics now render as three full-width rows in a fixed order. Each row reserves compact title and state cells while the path/detail cell receives all remaining horizontal space.
- Path/detail text stays on one line. Only text that still exceeds the available full-width row uses an ellipsis, and every detail cell carries its complete value in a hover tooltip.
- The layout is scoped to `#permissionsPane .permission-diagnostic-grid`, so dashboard and other diagnostic-card layouts are unchanged.
- Family Browser's canonical product, dashboard, assembly, and installer version is now `1.0.1`; all five packaged Revit targets report file/product version `1.0.1.0 / 1.0.1`.

### Regression Coverage

- Added `admin-permission-long-path-layout` at `1280x720` with long local, central, and Guard-root paths.
- The IE `WebBrowser` audit now requires exactly three diagnostic rows, full container width, strictly vertical ordering, a usable path cell, and a complete authored tooltip.
- The project-subtitle checker now follows each scenario's actual local/central path instead of assuming one hard-coded fixture path.
- Contract and workflow readers now request UTF-8 explicitly, preventing Windows PowerShell 5.1 from corrupting Korean audit tokens in BOM-less files.
- Final static UI/action/contract checks passed for all three hosts: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- Korean/English long-path rendering and common structured-result dialogs passed `24` executed IE `WebBrowser` scenarios with zero failures for Revit 2019/2023/2025/2027. Revit 2021 was `SKIP runtime-not-installed` while its shared staged assembly was verified.
- Workflow audit passed `970/970` checks across `35` workflows at `artifacts\family-browser-workflow-audit\20260721-v101-permission-layout-pwsh`.
- Visual reference: `artifacts\family-browser-ui-audit\20260721-v101-permission-layout-rerun-2019\permission-long-path-1280x720.png`.

### Build, ProgramData, And Installer

- Revit 2019/2021/2023/2025/2027 Release builds and the `17`-file Stage manifest passed integrity verification.
- ProgramData checkpoint: `_backups\programdata-before-v101-permission-layout-20260721-155759` (`17` files, `21,926,806` bytes).
- Elevated ProgramData deployment and installed-payload verification passed for all five targets. Final installed host DLL SHA-256 values are 2019/2021/2023 `7AAB7BD9129A7B55A918E2058C26359ABAEBA87CF344BD4A51ADF0439BD4FD64`, 2025 `8974E5FF4B87A9D50008ECEA75DEA3DEED5037C8B79097E9719B3D59D91CD651`, and 2027 `DE342C06602270780ED688DA8E177F7942478C53F701D023C640C74DF684E3F4`.
- Installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0.1_permission-layout-20260721-1558_Setup.exe`.
- Installer size: `3,821,807` bytes (`3.64 MB`). SHA-256: `8EB8FD9C14345E3BE1E3460D00581A87ED1E8609308BAF7D07C5931D8981C502`; the executable, `latest-build.json`, and companion `.sha256.txt` match.
- Existing installer artifacts were preserved and no mail-sized package was generated for this installer-only request.
- The local executable and ProgramData installation are `1.0.1`; no public website update feed was published or changed in this task.
- Revit was not launched automatically. `Needs User Revit Check`: restart the desired Revit version, open Permissions / Guard, and confirm real local/central/managed-folder paths remain readable in the three stacked rows and reveal their full text on hover.

## 2026-07-21 Version 1.0.1 Mail Package

Status: `Mail Package Complete / Distribution Verified`

- Reused the already verified `1.0.1` installer without rebuilding or changing its contents.
- Mail package: `artifacts\family-browser\mail-packages\20260721_01.zip`.
- Package size: `17,196,877` bytes (`16.40 MB`). SHA-256: `348C6BC5A46FBB6B5477FAE441F8F9BFF10B3170E3F118DA1C41B7FA91986A47`; the package, `latest-mail-package.json`, and companion `.sha256.txt` match.
- ZIP entries are `Setup.exe`, `README.txt`, and `mail_size_padding_do_not_run.bin`. The padding file is non-executable and exists only to meet large-file mail attachment handling.
- The ZIP-internal `Setup.exe` SHA-256 is `8EB8FD9C14345E3BE1E3460D00581A87ED1E8609308BAF7D07C5931D8981C502`, exactly matching the verified standalone `1.0.1` installer.
- Stage, standalone installer, mail-package contents, metadata, and installed ProgramData payload verification passed at `artifacts\family-browser-distribution-audit\permission-layout-20260721-1558-mail`.
- Revit was not launched and ProgramData did not require another deployment because the installer and Stage were not rebuilt during mail packaging.

## 2026-07-21 Central Synchronization Tracking Performance

Status: `Source/Stage/ProgramData Complete / Needs User Central Revit Performance Check`

### Cause

- A successful Save or Synchronize boundary always rebuilt the complete element-tracking baseline with `CaptureBaseline`, even though `DocumentChanged` had already maintained current state for the affected element IDs. On a large central model this repeated a full pass over trackable instances and element types after every synchronization.
- Every successful Save, Save As, and Synchronize also called `FamilyBrowserProjectCatalogService.Observe` unconditionally. That performed another complete Family/ElementType name catalog scan even when the commit only created or modified ordinary model instances such as walls.
- Immutable history/checkpoint persistence also remains synchronous by design. This preserves the existing audit durability boundary, but a slow management share can still contribute measurable latency independently of model comparison.

### Correction

- Source checkpoint: `_backups\20260721-sync-tracking-performance`. ProgramData checkpoint: `_backups\programdata-before-sync-tracking-performance-20260721-171255` (`17` files, `21,937,430` bytes).
- Normal successful commits now refresh only the union of locally pending IDs and incoming Sync/Reload IDs, then promote the already-maintained current state as the next baseline. Ordinary synchronization no longer recaptures the full model.
- The full-model baseline path remains a conservative fallback whenever changed-state refresh fails, an external rebase fails, a `DocumentChanged` read gap exists, or commit-boundary evidence is incomplete.
- Project catalog observation is now conditional. It runs when a Family or ElementType changed, when recovered local-save evidence contains such a change, or when tracking evidence is uncertain. Ordinary instance-only commits skip the full Family/ElementType catalog scan.
- Tracking coverage, Undo/Redo handling, local-save checkpoints, immutable history, incoming-central overlap attribution, and offline spool behavior were not removed or weakened.
- Added local timing diagnostics at `%LOCALAPPDATA%\KKY\FamilyBrowser\Diagnostics\element-tracking-performance.log`. Commit rows report total time, changed-state refresh, persistence, baseline mode, local/incoming ID counts, and whether a catalog scan was required. Separate `project-catalog` rows report whether that scan ran and its duration.
- Updated the workflow/static quality checks so they verify the conditional helper and incremental/full-fallback policy instead of requiring the previous unconditional call shape.

### Automated Verification

- Workflow audit passed `984/984` checks across `35` workflows at `artifacts\family-browser-workflow-audit\20260721-sync-tracking-performance-final`.
- Static UI and action-route checks passed for all three source hosts: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host. UI contract checks also passed.
- Revit 2019/2021/2023/2025/2027 Release builds completed and the `17`-file Stage manifest passed integrity verification. Existing unrelated 2025/2027 compiler warnings remain; this change introduced no build error.
- ProgramData installed-payload verification passed for all five versions. Stage and installed DLL SHA-256 values match exactly: 2019/2021/2023 `5FF87D02CEEF9F518E79422C5FA86938E73D9F0902ED3C2E3EB63F7FE2784F9D`; 2025 `75AADC1BD271C75A6EEAD902347D6389987E397FD7D0DD38A8C0CBE70EEB5F43`; 2027 `15AD830C543C5492E1571750C11D0BE73BC3262249F26CE58F65EC64AD61C317`.

### Runtime Boundary

- Revit was not launched automatically because the user performs runtime checks and automated Revit startup can trigger the reported licence issue.
- `Needs User Central Revit Performance Check`: restart Revit, open the same tracked central project, create or modify several ordinary instances, and synchronize. The corresponding commit log should show `baselineMode=incremental` and normally `catalogScan=skipped`; the following `project-catalog` row should show `performed=no`.
- If synchronization is still slow, compare `changedStateRefreshMs`, `persistenceMs`, `baselineMs`, and the `project-catalog` elapsed value. A dominant `persistenceMs` means the remaining delay is management-share I/O rather than model comparison; changing that boundary to background upload requires a separate durability design decision.

## 2026-07-21 File Guard Trade Assignment And Automatic First-Open Model Check

Status: `Source/Stage/ProgramData Complete / Needs User Revit And Two-PC Check`

### Root Cause And Intended Boundary

- File Guard identified which RVT received permissions and tracking, but its target record did not identify which separated-trade standard governed that project. Users therefore had to select the trade elsewhere before Current Model Check, and nested-only candidate loading could combine catalogs from unrelated trades.
- The File Guard workbook was export-oriented and had no `공종` / `Discipline` round-trip. A bulk policy could not express the project-to-standard relationship required for automatic checking.
- Automatic checking must apply only to an exact File Guard target, must not rescan an unchanged project, must not create an Excel workbook by itself, and must not let two Revit sessions publish competing comparison results for the same project.

### Correction

- Source checkpoint: `_backups\file-trade-auto-model-check-20260721-174639`.
- Added `Discipline` to every Revit host's `FamilyBrowserFileGuardTarget`. The active HTML File Guard window now has a trade selector on every RVT row, a bulk visible-row trade assignment, Excel import, and Excel template/current-policy export.
- Korean workbooks contain `공종`; English workbooks contain `Discipline`. Import accepts either language plus common aliases, normalizes the value to a currently registered standard slot, preserves every existing per-file guard flag, warns on blank legacy trade fallback, skips invalid/unregistered trades, and replaces duplicate RVT paths deterministically.
- Saving the form requires an assigned registered trade for every enabled guarded RVT. Existing policies without the new field remain readable; the UI supplies the current effective trade so saving converts them to an explicit assignment.
- Added `FamilyBrowserAutomaticModelCheckService`. `DocumentOpened` schedules the check, then safe Revit Idling passes wait for document activation and a nonmodifiable state. The service resolves the exact File Guard target and its assigned trade, validates the registered/scanned standard, restores a revision-matching saved project comparison when possible, and captures/builds/saves a new comparison only when the cache is absent or stale.
- A project-scoped `automatic-model-check.lock` prevents simultaneous sessions from scanning and publishing the same project comparison. A waiting session rechecks and reuses the result produced by the first session. No modal result window or XLSX is created automatically; the dashboard Model Check pane shows the assigned target and current status.
- Nested-only standalone-placement fingerprint refresh/evaluation now receives the same per-file trade. Candidate and evidence cache keys include that trade/slot, so a child that is nested-only in another discipline no longer governs the current project's placement policy.
- Revit was not launched during this work because automated startup on this PC can trigger the known licence error.

### Automated Verification

- Static UI/action/contract checks passed for all three host source sets: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- Workflow audit passed `984/984` checks across `35` workflows at `artifacts\family-browser-workflow-audit\20260721-181303`.
- Dedicated nested-family difference propagation and authoritative System Type apply checks passed.
- Revit 2019/2021/2023/2025/2027 Release compilation completed with zero errors, and the `17`-file Stage manifest passed integrity verification.
- Focused Korean/English IE `WebBrowser` checks passed `36/36` scenarios for the 2019, 2025, and 2027 representative hosts at `artifacts\family-browser-ui-audit\20260721-file-trade-auto-check-targeted`. An earlier broad run completed `80` scenarios with zero failures before the outer 10-minute orchestration limit, so it was not counted as a complete full-suite pass.
- The compiled 2019 host then passed a real Korean XLSX File Guard export/import round-trip at `artifacts\family-browser-ui-audit\20260721-file-trade-excel-roundtrip`: one RVT returned with `Mechanical` and all three tested guard flags unchanged.

### ProgramData Deployment

- ProgramData checkpoint: `_backups\programdata-before-file-trade-auto-model-check-20260721-182931` (`17` installed files plus `backup-info.txt`).
- Elevated deployment and installed-payload verification passed for Revit 2019/2021/2023/2025/2027. Every installed host reports file/product version `1.0.1.0 / 1.0.1`.
- Stage and installed host DLL SHA-256 values match exactly: 2019/2021/2023 `44EFEBC0329BC235BABDC9EE17A4AF3F93425205F1C9D98485FED61098AE2B96`; 2025 `4B94115D0D2B79A42FC1D3EEEE2AFB713F3812D5DCFA2EB64365F0495C45A0AB`; 2027 `B0FB1597C1FA96313966FA76860C802B400D4ABA6780262686BB23855B3AEA07`.

### Runtime Acceptance Check

- `Needs User Revit Check`: configure Architecture and Mechanical standards through registration, precise scan, and approved-list connection. Register two disposable project RVTs in Permissions / Guard and assign one trade to each, including one assignment through XLSX import.
- Restart/open each project. The Model Check pane must show the assigned trade. A stale or missing project cache must run one automatic comparison; an immediate unchanged reopen must report cached-result reuse. Changing the project or accepted standard revision must cause exactly one new check.
- Confirm no diagnostic/result workbook appears merely from the automatic run. XLSX must be created only after an explicit user export action.
- `Needs User Two-PC Check`: open the same guarded workshared project nearly simultaneously from two add-in-equipped PCs. One client must hold the scan lock; the second must wait and reuse the completed comparison without a duplicate published result.
- With a nested-only child that exists in a different trade, confirm Admin OFF direct-placement blocking follows only the project RVT's assigned trade. Revit 2021 remains `SKIP runtime-not-installed` on this PC, while its shared assembly, Stage, and ProgramData payload are verified.

## 2026-07-21 File Guard Trade Preservation Final Audit

Status: `Source/Stage/ProgramData Complete / Needs User Revit And Two-PC Check`

### Additional Finding And Correction

- Supplemental source checkpoint: `_backups\file-trade-preservation-final-20260721-185407`.
- The final all-constructor audit found that `FamilyBrowserStandardPolicyStore.CloneFileGuardTarget(...)` copied every File Guard field except `Discipline`. A policy mutation that cloned the current policy could therefore silently clear the RVT's assigned trade even though the HTML form and Excel import had saved it correctly.
- All three host source sets now copy `Discipline` during policy cloning and trim it during policy normalization. Protected dashboard audit fixtures also carry the active trade so the automated scenarios exercise the new contract instead of a legacy blank assignment.
- The compiled UI harness now calls the private policy clone path through reflection and fails if `Mechanical` is lost. Static checks also require both clone preservation and normalization in every host.

### Final Verification And Deployment

- Static UI/action/contract checks passed for all three hosts: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- Workflow audit passed `984/984` checks across `35` workflows at `artifacts\family-browser-workflow-audit\20260721-file-trade-preservation-final`. A Windows PowerShell 5.1 invocation first produced Korean-token mojibake; rerunning the same UTF-8 source under the workspace PowerShell 7 runtime passed and confirmed that it was an audit-runner encoding issue, not a product failure.
- The compiled Korean/English IE `WebBrowser` audit passed `18/18` focused scenarios with zero failures for the 2019, 2025, and 2027 representative hosts at `artifacts\family-browser-ui-audit\20260721-file-trade-preservation-final`. The 2019 assembly is shared by 2019/2021/2023.
- Dedicated nested-family difference propagation and authoritative System Type apply checks passed.
- Revit 2019/2021/2023/2025/2027 Release builds and the `17`-file Stage integrity verification passed.
- Final ProgramData checkpoint: `_backups\programdata-before-file-trade-preservation-final-20260721-190100` (`17` files, `22,320,710` bytes).
- Elevated ProgramData deployment and installed-payload verification passed for all five targets. Stage and installed DLL SHA-256 values match exactly: 2019/2021/2023 `1BBE91A8691F34CD2915C3C3528361A4EDE74E797F45D7634665852EE2EEF7EE`; 2025 `C6ABB7B36E1B6487187C54C8D11029725A17FAF1D14454CF04E6957A531A0547`; 2027 `C22CABCD42824B3B87C85CA5FD4E10A505389B469FCF22AE80DB2079B1710BD2`.
- Every installed host reports file/product version `1.0.1.0 / 1.0.1`. Revit process count remained `0`; Revit was not launched.

### Remaining Runtime Boundary

- `Needs User Revit Check`: assign different trades to two guarded RVTs, close and reopen the File Guard window, switch another policy setting that invokes a policy mutation, and confirm both assignments remain unchanged without Refresh.
- Open each guarded RVT and confirm the first safe idle automatically checks only its assigned trade standard. Reopen an unchanged file and confirm the cached result is reused rather than rescanned.
- `Needs User Two-PC Check`: open the same guarded workshared RVT from two equipped PCs and confirm one automatic check publishes while the other waits and reuses it.

## 2026-07-21 File Guard Trade / Automatic Check Pre-Deployment Review

Status: `Review Complete / Source Fix Required / Revit Not Launched`

### Confirmed Findings

- `P1 - UNC endpoint overmatching`: `FamilyBrowserFileGuardPathMatcher.BuildEndpointNeutralUncKey(...)` removes the server/IP segment completely. Two unrelated paths such as `\\SERVER-A\\bim\\Project.rvt` and `\\SERVER-B\\bim\\Project.rvt` are therefore considered identical whenever the share-relative path matches. Hostname/IP alias support must prove endpoint equivalence instead of ignoring the endpoint.
- `P1 - Alias duplicates can disable the guard`: File Guard HTML rows and XLSX import deduplicate with `Path.GetFullPath(...)` text only, while runtime matching uses mapped-drive/UNC/file identity. The same RVT can be stored twice through aliases; runtime then returns `AmbiguousPathIdentity` with no target, so permissions, tracking, automatic checking, and nested-only blocking can be skipped.
- `P1 - Project cache freshness is not the current document freshness`: `ProjectSnapshotStore` stamps only `ResolveProjectIdentityPath(...)`, which is the central path for workshared documents. Unsynchronized local changes are invisible to this stamp. Cloud/Model Server identities that cannot be read through `FileInfo` have no project freshness stamp at all, so an old comparison can be reused after the current document changed.
- `P1 - Two-PC wait can expire in seconds`: scan-lock contention retries after three Idling ticks and stops after twelve attempts. Revit Idling is not a time interval; a second client can become `Deferred` long before the first precise scan finishes. On the first scan, mapped-drive/UNC/IP aliases can also produce different project-history keys and therefore different lock files, allowing both sessions to scan and publish concurrently.
- `P1 - Nested-only placement blocks before an exact fingerprint match`: `PendingVerification` and `VerificationUnavailable` are converted to blocked native change events just like `ExactMatch`. This is stricter than the requested rule, which blocks only the same name/category/shared state and same precise fingerprint. A different family can be rejected temporarily or indefinitely when verification fails.
- `P2 - Legacy blank trade is nondeterministic`: automatic checking and nested-only runtime resolution call `ResolveAssignedDiscipline(..., allowLegacyFallback: true)`. A legacy RVT without an explicit trade silently follows whichever global standard trade is currently effective instead of requiring migration.
- `P2 - Blank XLSX trade silently adopts current UI state`: an empty `공종/Discipline` cell imports as the current effective standard trade. Bulk import should reject or explicitly migrate blank rows so the same workbook produces the same policy regardless of the administrator's current selection.
- `P2 - Automatic check can freeze the Revit UI`: the stale-cache path runs precise project capture, comparison, and managed-folder persistence synchronously inside the application Idling callback. There is no defer/cancel decision, and progress-message updates cannot repaint while the same UI thread is occupied.
- `P2 - Non-success automatic outcomes are not persisted`: missing managed folder, missing trade/standard/scan, lock deferral, and caught failures update only the in-memory request. `automatic-model-check-latest.json` is written only for cache reuse or success, so an administrator cannot audit why a first-open check did not run after the session closes.

### Verification Gap

- `Test-FamilyBrowserUiStatic.ps1`, `Test-FamilyBrowserUiContract.ps1`, and `Test-FamilyBrowserWorkflow.ps1` all passed on this review run. The workflow result is `artifacts\\family-browser-workflow-audit\\20260721-193655`.
- Current tests assert that hostname and IP strings with the same share-relative path must match, but do not reject two genuinely different servers with the same share-relative path.
- Current tests check that the new methods and routes exist; they do not simulate elapsed lock contention, a workshared local with unsynchronized changes, a non-file cloud identity, blank legacy trade migration, or pending/failed nested fingerprint behavior.

### Required Runtime Boundary

- Do not treat the latest ProgramData build as final for this feature until the P1 items above are corrected and rebuilt.
- Revit remains closed because automated launch on this PC can trigger the known licence error.
- After source fixes, run one mapped-drive/UNC/IP same-file case, one deliberately different-server same-relative-path case, one unsynchronized workshared-local case, one two-PC slow-scan case, and exact/different/failed nested-fingerprint placements.

## 2026-07-21 Critical File Guard / Automatic Check Remediation

Status: `Source/Stage/ProgramData Complete / Automated Regression Passed / Needs User Revit And Two-PC Check`

### Source Checkpoint

- Source checkpoint: `_backups\critical-review-remediation-20260721-194731`.
- ProgramData checkpoint: `_backups\programdata-before-critical-remediation-20260721`.
- Revit was not launched because automated startup on this PC can trigger the known licence error.

### Corrected Findings

- `P1 - UNC endpoint overmatching`: removed endpoint-neutral `UNC-RELATIVE` matching. Hostname/IP strings are no longer treated as the same server merely because their share-relative text matches. Exact normalized paths, mapped-drive resolution, Windows physical file identity, or the existing explicit unique-workshared-name startup fallback are required.
- `P1 - Alias duplicates can disable the guard`: File Guard HTML and XLSX paths now use `BuildStablePolicyPathKey(...)`, based on the same comparable/physical identity used at runtime. If legacy rows still resolve to the same physical RVT, runtime creates a conservative merged target: every restriction is ORed, tracking is preserved, and conflicting trade assignments require explicit correction instead of silently disabling the guard.
- `P1 - Project cache freshness`: project-cache schema is now `3` and records `ProjectDocumentRevisionToken`. Live documents reject cache reuse while `Document.IsModified` is true, when the token is absent, or when the Revit basic-file/central episode revision differs. This prevents unsynchronized in-memory work from reusing an older central-path-only comparison.
- `P1 - Two-PC coordination`: lock files now use `GetProjectCoordinationFolder(...)`, derived from stable physical project identity rather than whichever mapped/UNC alias a client opened. Contention now retries by elapsed time for up to 30 minutes at five-second intervals instead of expiring after a few Idling ticks.
- `P1 - Nested-only placement`: native rollback is now emitted only for `ExactMatch`. `PendingVerification` and `VerificationUnavailable` remain nonblocking and are not reported as a prohibited placement.
- `P2 - Legacy/Excel trade fallback`: automatic checking and nested-only refresh use `allowLegacyFallback: false`. Blank or invalid XLSX `공종/Discipline` rows are skipped with an explicit warning; they never inherit the administrator's currently selected trade.
- `P2 - Automatic-check diagnostics`: Scheduled, Running, Waiting, Deferred, configuration failures, caught failures, cache reuse, and success are persisted to `automatic-model-check-latest.json`. Running progress fields were added to status schema `2`.
- `P2 - Silent UI freeze`: stale-cache Revit API capture remains on the required Revit UI context, but now opens a dedicated noninteractive progress surface and synchronously repaints it without `Application.DoEvents`. Capture, snapshot save, comparison, and publication occupy stable progress ranges; cache reuse still bypasses the window.

### Automated Verification

- Static UI/action/contract checks passed for all three host source sets: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- Nested-family difference propagation and authoritative System Type apply checks passed.
- Workflow audit passed at `artifacts\family-browser-workflow-audit\20260721-critical-remediation`; the quality-gate workflow copy also passed at `artifacts\family-browser-ui-audit\20260721-critical-remediation-quality-gate\workflow`.
- The 2,000-row cache/filter performance gate passed with zero failures at `artifacts\family-browser-ui-audit\20260721-critical-remediation-quality-gate\performance`.
- Revit 2019 full IE `WebBrowser` coverage completed `58` Korean/English scenarios with zero failures. Revit 2025 and 2027 each passed `14` focused Korean/English scenarios covering messages, Home, Admin OFF File Guard, Family detail/preview, System detail, and Standards layout at `artifacts\family-browser-ui-audit\20260721-critical-remediation-targeted-2025` and `...-2027`.
- Revit 2021 runtime remains `SKIP runtime-not-installed`. Revit 2023 uses the same verified 2019/2021/2023 host assembly and its structured-message smoke passed.
- The broad quality-gate wrapper reached its external ten-minute limit during redundant multi-version screenshot generation after static, workflow, five-version build/Stage, Stage verification, and performance steps had already passed. No generated result JSON contains a failure, and no harness process was left running.
- Revit 2019/2021/2023/2025/2027 Release builds completed with zero errors. The `17`-file Stage manifest and every staged add-in manifest passed integrity verification.

### ProgramData Deployment

- The existing five-version ProgramData payload was backed up at `_backups\programdata-before-critical-remediation-20260721` before deployment.
- The elevated installer was released from an accidental Windows console selection pause and completed normally. Revit 2019/2021/2023/2025/2027 are all updated in ProgramData.
- `Verify-FamilyBrowserRecovered.ps1 -Installed -Years 2019,2021,2023,2025,2027` passed. The captured verification output is `artifacts\family-browser-ui-audit\20260721-critical-remediation-installed-verify.txt`.
- Every staged file matches its ProgramData counterpart: 2019/2021/2023 each checked `3` files with zero mismatches; 2025/2027 each checked `4` files with zero mismatches.
- Stage and installed host DLL SHA-256 values match exactly: 2019/2021/2023 `F105DF948E38A33AF8B14D1C6F20E216243FBB42C7F74AF2397126D8EA55FBEC`; 2025 `61338E960EA10571CACE71645116E0B2E1D46E619B84ACE0E84BD5267574FA14`; 2027 `F935E900DBF08DEAC1AA5C9D7D21C304A7345CC66B370CCBBE74A9A08F931064`.
- No Revit, installer, or UI-audit process was left running.

### Remaining Runtime Boundary

- `Needs User Network Revit Check`: verify one real mapped-drive/UNC alias of the same RVT and one deliberately different server with the same share-relative path. Only the physically proven alias may inherit the guard.
- `Needs User Unsaved Revit Check`: modify a guarded workshared local without save/sync and confirm automatic checking does not reuse the previous project cache; save or synchronize, reopen unchanged, and confirm cache reuse.
- `Needs User Two-PC Check`: hold a precise scan on one equipped PC longer than the former retry window and open the same RVT through another alias on a second PC. The second client must wait and reuse the first publication without creating a competing result.
- `Needs User Nested Placement Check`: test exact fingerprint, different fingerprint, and unavailable fingerprint cases. Only the exact standard match may roll back standalone placement while Admin Mode is OFF.
- `Needs User Progress Check`: force one stale automatic first-open check and confirm the progress window paints through capture/save/compare/publication without permitting another Revit command to be posted.

## 2026-07-21 Critical Remediation Second-Pass Review

Status: `Review Complete / Source Fix Required / Current ProgramData Not Final / Revit Not Launched`

### Confirmed P1 Findings

- `Active document switch does not refresh an already-open dashboard`: all three current `FamilyBrowserDashboardModelessRuntime.NotifyActiveDocumentChanged(...)` implementations update `_lastAutoDocumentKey` but never retain `_form` or call `RefreshForActiveDocumentChanged(...)`. The form refresh method still exists and the recovery copy contains the missing call, so switching between open RVTs can leave document title, permissions, rows, automatic-check status, and cached comparison data from the previous project on screen.
- `A dirty in-memory project can be published as the shared latest comparison`: `SaveLatestProjectScan(...)` reads `Document.IsModified` but discards it. The cache record has no writer-dirty field, and a workshared local whose on-disk `BasicFileInfo` still says all local changes are synchronized receives the same central revision token as a clean client. Manual and automatic checks can therefore publish unsaved local content that another clean client later accepts as the shared latest result.
- `Automatic first-open checking does not validate the live Standard RVT revision`: `FamilyBrowserAutomaticModelCheckService.Execute(...)` loads the registration and stored snapshot, then reuses or creates a project comparison without calling `FamilyBrowserStandardRevisionService.Probe(...)` before capture or after capture. A changed Standard RVT can therefore be compared through its old snapshot and the project can be marked checked against stale standard data.
- `Startup preload can reuse stale comparison data when the project was edited before Family Browser opened`: background preload receives only `FamilyBrowserDeploymentProjectIdentity`, so it bypasses live `Document.IsModified` and document-revision checks. The later guard uses only mutations observed after the form was created; edits made before opening the browser leave its serial at zero and can accept the prepared stale result.
- `Transient mapped-drive/UNC identity failure can fail open for both Guard and tracking`: empty mapped-drive resolution is cached for 30 seconds. When a workshared document exposes a usable mapped/UNC string but alias equivalence cannot temporarily be proven, filename fallback is disabled and no target is returned. Native Guard then permits protected commands, while element tracking can treat the same result as a definite out-of-scope policy and end a session instead of marking identity resolution deferred.
- `Duplicate physical File Guard rows can be weakened by the configuration UI or XLSX import`: runtime matching conservatively OR-merges duplicate physical targets, but the HTML form keeps the first matching row and the Excel importer replaces it with the last row. Opening/saving a legacy alias-duplicate policy or importing conflicting duplicate rows can therefore lose stricter block/tracking flags or disable the target.
- `Project scan/history publication is not atomic and manual scans bypass the coordination lock`: project snapshots, comparison reports, latest-cache records, and aliases use direct `File.WriteAllText(...)`; history filenames have only second precision. Automatic checking owns `automatic-model-check.lock`, but the manual Current Model Check publishes through the same paths without that lock. Two clients or manual/automatic overlap can overwrite same-second history, expose partial JSON, or publish mismatched latest pointers.
- `Nested-only standalone placement has a verification-window bypass`: the current contract intentionally blocks only `ExactMatch`; `PendingVerification` and `VerificationUnavailable` are allowed. Fingerprints are prepared one family per Idling pass and no later rollback or violation record examines placements made during that window. The requested “same standard nested-only family cannot be placed standalone” guarantee is therefore not absolute until verification is prewarmed or pending placements are rechecked.

### Confirmed P2 Findings

- `Manual Current Model Check can use a Standard RVT state cached for up to the 55/60-second probe interval`: the dashboard checks its cached revision state before a deep scan but does not synchronously probe and revalidate the same revision after capture. This can publish a report across a Standard RVT change that occurs just before or during the scan.
- `Project alias recovery can select a same-name collision`: non-live startup identity accepts an alias record using project name plus file timestamp and length. Different projects with the same basename, size, and timestamp can resolve to the wrong cache before a live-document revision check is applied.
- `The V2 manifest has a cross-process lost-update race`: Standard artifact, approved-list, and project-state publishers atomically replace individual files, but each independently performs an unlocked manifest read-modify-write. Concurrent publishers can restore stale manifest fields and erase another publisher's pointer/revision update.
- `Physical file identity is not namespaced by remote endpoint`: `GetFileIdentity(...)` stores only volume serial and file index. For remote files, an identity key should also be scoped by a proven canonical endpoint/final-path namespace to avoid theoretical collisions across separate servers.

### Automated Verification And Coverage Gap

- `Test-FamilyBrowserUiStatic.ps1` passed for all three hosts: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- `Test-FamilyBrowserUiContract.ps1` passed. The second-pass output is `artifacts\family-browser-ui-audit\20260721-second-pass-review-contract`.
- `Test-FamilyBrowserWorkflow.ps1` passed `984/984` checks across `35` workflows at `artifacts\family-browser-workflow-audit\20260721-second-pass-review`.
- These passes do not clear the findings. Existing tests assert that the relevant methods/tokens exist, that a dirty reader rejects cache reuse, and that pending nested fingerprints remain nonblocking. They do not execute an active-RVT switch, a dirty writer followed by a clean reader, a Standard RVT revision change during automatic/manual capture, conflicting UI/XLSX aliases, mapped-drive resolution failure, manual/automatic publication overlap, or unlocked manifest writers.

### Required Fix Order

1. Restore active-document dashboard refresh and add a document-switch regression test.
2. Prevent dirty documents from shared publication, record writer state in the cache schema, and validate live project and Standard RVT revisions both before and after capture.
3. Make mapped/UNC identity uncertainty explicit and conservative for already-targeted open documents; merge or reject duplicate UI/XLSX policies without weakening flags.
4. Put every manual/automatic project publication behind one project-scoped lock and use atomic promotion plus collision-free history filenames; serialize manifest mutations.
5. Close the nested-only verification window through prewarming or post-verification violation handling, then add Revit runtime acceptance cases.

### Runtime Boundary

- Revit was not launched because automated startup on this PC can trigger the known licence error.
- No product source, Stage, installer, or ProgramData payload was changed during this review. The currently installed build still contains the findings above and must not be treated as the final deployment candidate.

## 2026-07-21 Critical Remediation Second-Pass Completion

Status: `Source/Stage/ProgramData Complete / Automated Regression Passed / Needs User Revit And Two-PC Check`

### Checkpoints

- Main source checkpoint: `_backups\critical-second-pass-remediation-20260721-211655`.
- Installer-tool checkpoint: `_backups\installer-write-access-probe-20260721-222552`.
- ProgramData checkpoint before deployment: `_backups\programdata-family-browser-before-critical-second-pass-20260721-222455` (`17` files, `22,376,942` bytes).
- Revit was not launched. The user performs runtime checks because automated startup on this PC can trigger the reported licence error.

### Corrected Findings

- `Active-document refresh`: every modeless runtime now retains the open dashboard and calls `RefreshForActiveDocumentChanged(...)` outside its synchronization lock. Switching active RVTs can no longer update only the internal key while leaving the previous project's title, policy, rows, and automatic-check status on screen.
- `Dirty project publication/cache reuse`: project-cache schema is now `4` and records `CapturedFromModifiedDocument`. Manual and automatic scans reject modified documents both before capture and before publication; readers reject older schema and dirty-writer records. Startup preload must revalidate through the live `Document`, so edits made before opening Family Browser cannot accept a prepared stale result.
- `Standard RVT revision boundary`: manual and automatic model checks force a live Standard RVT revision probe before capture and validate the same revision again after capture. A changed, unavailable, or stale Standard RVT aborts publication instead of marking the project checked against an old snapshot.
- `Transient mapped/UNC identity`: workshared central-path identity uncertainty is explicit. A unique registered filename remains conservatively guarded while identity is unavailable; multiple same-name candidates produce one strict OR-merged temporary target instead of failing open. A resolved different central path still receives no filename fallback.
- `Duplicate File Guard policy`: HTML configuration and XLSX import now merge duplicate physical aliases with the same conservative policy rule. Restrictive family/type/nested-only/tracking flags cannot be weakened by first-row/last-row ordering, and conflicting trade values remain blank for explicit administrator resolution. UI row identity prefers a stable physical/path key rather than filename alone.
- `Atomic project publication`: manual and automatic project scans share one project-scoped publication lock. Snapshot, comparison, latest, alias, and history JSON writes use atomic promotion; history names include high-resolution UTC plus a GUID so same-second writers cannot collide.
- `Manifest concurrency`: V2 manifest read-modify-write operations now hold a cross-process mutation lock, preventing Standard artifacts, approved-list updates, and project-state publication from erasing one another's fields.
- `Alias recovery`: project aliases are accepted only when the recorded and requested comparable identities match, closing the same-basename/same-length/same-timestamp collision path.
- `Nested-only verification window`: a verified project snapshot now seeds exact nested-only fingerprints by Family `UniqueId` before normal Idling verification. Already-scanned exact matches are protected immediately on reopen/automatic/manual check, while pending or unavailable evidence remains nonblocking to avoid false rollback.
- `Installer ACL handling`: the ProgramData installer no longer assumes that every non-elevated process is unwritable. It probes the exact add-in/payload folders and existing manifest access before mutation, while still failing before changes when the ACL is insufficient. This PC required the elevated path because existing DLL/manifest files are administrator-owned.

### Automated Verification

- Static UI/action/contract checks passed for all three host source sets: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- Workflow audit passed `984/984` checks across `35` workflows at `artifacts\family-browser-ui-audit\20260721-critical-second-pass-quality-final\workflow`.
- Dedicated nested-family difference propagation and authoritative System Type apply tests passed.
- The 2,000-row performance/cache gate passed with zero failures at `artifacts\family-browser-ui-audit\20260721-critical-second-pass-quality-final\performance`. Across installed runtime families, shell generation was `1-2 ms`, 1,000-row cold rendering was `199-364 ms`, warm rendering was `36-63 ms`, filtering was `10-16 ms`, and only `150` rows were present in the DOM at once.
- IE `WebBrowser` HTML/click coverage completed every unique Korean/English scenario with zero failures for Revit 2019, 2023, 2025, and 2027. Revit 2019, 2025, and 2027 each completed `58/58`; the initial broad run produced `41` Revit 2023 results before its external 15-minute limit, and `artifacts\family-browser-ui-audit\20260721-critical-second-pass-harness-split\Rvt2023-missing` completed every remaining unique case. Revit 2021 is `SKIP runtime-not-installed`; its shared 2019/2021/2023 binary and Stage payload are verified.
- The broad quality wrapper's first process reached its external 15-minute limit during redundant multi-version screenshot/click generation after static, contract, workflow, managed-data, five-version build, Stage verification, and performance checks had passed. Its `99` completed result JSON files contained zero failures; the split runs then completed the remaining unique host scenarios.
- Managed-data audit had zero failures and one expected environment warning at `artifacts\family-browser-ui-audit\20260721-critical-second-pass-quality-final\managed-data`: homepage bootstrap returned HTTP 200, the unavailable `I:` candidate was reported, and reachable `D:\TEST` was selected. Only Revit 2019 currently contains managed-data folders on this PC.
- Revit 2019/2021/2023/2025/2027 Release builds completed with zero errors. The only compiler warning is the existing `NETSDK1137` SDK-style notice for the 2025/2027 projects. The `17`-file Stage manifest passed integrity verification.

### ProgramData Deployment

- Elevated ProgramData deployment completed for Revit 2019/2021/2023/2025/2027. Installed-payload verification is captured at `artifacts\family-browser-ui-audit\20260721-critical-second-pass-installed-verify.txt`.
- All `17` staged files match their ProgramData counterparts by SHA-256; mismatch count is `0`.
- Installed file/product version remains `1.0.1.0 / 1.0.1` as requested. Host DLL SHA-256 values are: 2019/2021/2023 `8338795AA6E0EFA49ACCE286DA57BE0AA90349C50C8034D616A102B99124BA08`; 2025 `095A09284A07C8B6848926A12FB52F3E073ED080B41A5A35AB2EBE18F796CFEA`; 2027 `88608CFCC15095DA460B67BEE98FDEC762AAA185E45BF876D946912CD9B27415`.

### Remaining Runtime Boundary

- `Needs User Active-RVT Check`: keep Family Browser open, switch between two RVTs with different File Guard trades/policies, and confirm document information, rows, permissions, automatic-check status, and detached detail content refresh immediately without pressing Refresh.
- `Needs User Dirty/Revision Check`: modify a guarded project without save/sync and confirm manual/automatic checks refuse shared publication. Change the Standard RVT immediately before and once during a test scan; neither run may publish against the previous revision.
- `Needs User Network Check`: temporarily disconnect and reconnect the mapped/UNC management path. A previously registered workshared target must remain conservatively blocked during identity uncertainty, while a resolved different central RVT with the same filename must not inherit that policy.
- `Needs User Two-PC Check`: overlap manual and automatic checks from two equipped clients against the same RVT and confirm one project publication is serialized and the second reuses the completed result. Also overlap Standard/list/project-state publishers and confirm the V2 manifest retains all pointers.
- `Needs User Nested Placement Check`: after a verified project snapshot, directly place an exact nested-only child and confirm rollback while Admin Mode is OFF; different or unavailable fingerprints must remain allowed. A first-ever placement before any precise evidence remains intentionally nonblocking and is a policy/design boundary rather than a false-positive block.
- `Needs Design / very low probability`: Windows remote file identity still exposes volume serial plus file index without a provable canonical server namespace. Matching also requires the same RVT filename and is used only after textual path checks, but a theoretical cross-server ID collision cannot be eliminated safely without real DFS/hostname/IP fixtures.

## 2026-07-22 Critical Remediation Third-Pass Completion

Status: `Source/Stage/ProgramData Complete / Automated Regression Passed / Needs User Revit And Two-PC Check`

### Checkpoints

- Main remediation checkpoint: `_backups\critical-third-pass-remediation-20260721-225303`.
- Standard publication checkpoint: `_backups\critical-third-pass-standard-publication-20260721-233815`.
- Remote Standard refresh checkpoint: `_backups\critical-third-pass-remote-standard-refresh-20260722-0015`.
- Dirty-marker concurrency checkpoint: `_backups\critical-third-pass-dirty-marker-concurrency-20260722-0025`.
- Unique report-history checkpoint: `_backups\critical-third-pass-unique-report-history-20260722-0040`.
- Thumbnail metadata checkpoint: `_backups\critical-third-pass-thumbnail-atomic-20260722-000915`.
- Diagnostics checkpoint: `_backups\critical-third-pass-diagnostics-atomic-20260722-001523`.
- Automatic-status checkpoint: `_backups\critical-third-pass-automatic-status-atomic-20260722-001844`.
- ProgramData checkpoint before deployment: `_backups\programdata-family-browser-before-third-pass-20260722-004526` (`17` files).
- Revit was not launched. The user performs runtime checks because automated startup on this PC can trigger the reported licence error.

### Corrected Findings

- `Standard revision identity`: accepted Standard registration state now includes the snapshot path and snapshot timestamp. A re-scan of the same RVT path can no longer look unchanged merely because the selected file path stayed the same.
- `Remote same-source refresh`: Standard metadata/cache loading now detects a newer publication from another client even when the source path is unchanged, invalidates stale prepared data, and refreshes against the new revision.
- `Coordinated Standard publication`: Standard snapshot, derived browser artifacts, registration state, and manifest mutation are serialized and atomically promoted. Readers cannot observe an accepted revision whose related files are only partly published.
- `Dirty-marker concurrency`: concurrent project-dirty writers merge instead of replacing one another. A completed scan clears only the generation it captured, so a newer mutation arriving during capture is not erased.
- `Shared report history`: family scan, system scan, preflight, tracking, diagnostics, and error outputs use high-resolution UTC plus a GUID and atomic promotion. Same-second clients no longer overwrite one another or expose partially written JSON/text.
- `Thumbnail metadata`: thumbnail cache metadata is written through a sibling temporary file, flushed to disk, and atomically promoted. Interrupted or overlapping preview updates cannot leave truncated metadata behind.
- `Automatic-check status`: latest automatic-check status now uses UTF-8 without BOM, flush-to-disk, atomic promotion, and deterministic temporary-file cleanup.
- `Dialog cancellation`: the affected result/selection dialog now has its Cancel button wired to the form's `CancelButton`, restoring immediate Esc cancellation instead of leaving an inconsistent close path.

### Automated Verification

- Static UI/action checks passed for all three host source sets: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- UI contract checks passed for all three host source sets.
- Workflow audit passed `1037` checks. Latest result: `artifacts\family-browser-ui-audit\20260722-third-pass-diagnostics\workflow`.
- Dedicated nested-family difference propagation and authoritative System Type apply tests passed.
- IE `WebBrowser` HTML/click coverage passed every scenario for the three independent host implementations: Revit 2019 `58/58`, Revit 2025 `58/58`, Revit 2027 `58/58`, for `174/174` total. The 2019 assembly is shared by Revit 2019/2021/2023. Results: `artifacts\family-browser-ui-audit\20260722-third-pass-quality\harness-Rvt2019`, `...\harness-Rvt2025`, and `...\harness-Rvt2027`.
- The harness covered Korean/English, administrator/modeler, Admin OFF protected-family state, missing Standard RVT/list, unavailable/test managed folder, family/system details, component comparison, Standard settings, Model Check, requests/permissions, long paths, pending save/sync, and 1280x720, 1600x900, and 1920x1080 layouts.
- The 2,000-row performance/cache gate passed at `artifacts\family-browser-ui-audit\20260722-third-pass-quality\performance`. All available hosts stayed within the acceptance thresholds: shell generation `1 ms`, cold usable list `425-673 ms`, warm list `286-378 ms`, filter `10-11 ms`, and `150` rows in the DOM at once.
- Managed-data audit completed with zero failures and one expected environment warning at `artifacts\family-browser-ui-audit\20260722-third-pass-quality\managed-data`: homepage bootstrap returned HTTP 200, unavailable `I:` was reported, and reachable `D:\TEST` was selected. Only Revit 2019 currently has managed-data folders on this PC.
- Revit 2019/2021/2023/2025/2027 Release builds completed with zero errors. The large warning count consists of recovered/decompiled-source nullable/unused-field notices and Windows-only API analyzer warnings; no `CS1717` or unreachable-code `CS0162` warning was present. Build log: `artifacts\family-browser-ui-audit\20260722-third-pass-build\build.log`.
- The `17`-file Stage manifest and every staged add-in manifest passed integrity verification.

### ProgramData Deployment

- Elevated deployment completed for Revit 2019/2021/2023/2025/2027. Install log: `artifacts\family-browser-ui-audit\20260722-third-pass-quality\programdata-install.log`.
- `Verify-FamilyBrowserRecovered.ps1 -Installed` passed for all five targets.
- Every staged file matches its ProgramData counterpart by SHA-256: `17` checked, `0` missing, `0` mismatched.
- Installed version remains `1.0.1`. Host DLL SHA-256 values are: 2019/2021/2023 `F6BBAC8EDA3D93908298D837019224BFDB43B18314C1B991A0F303369BAAB8F2`; 2025 `0AE2BFF50966AA6F0D309B30357BFFEA1F150CFB03C981A039B4060677064E75`; 2027 `C3F962357EF53E92279F78354EC4155533182DF969272F467DB9390FE02F12F4`.

### Remaining Runtime Boundary

- `Needs User Active-RVT Check`: keep Family Browser open while switching between two RVTs with different policies/trades and confirm document information, list rows, permissions, automatic-check status, and detached detail content refresh without pressing Refresh.
- `Needs User Guard Check`: with Admin Mode OFF, verify immediate native Family Load blocking, immediate type/family rename rollback, mapped-drive/UNC/IP aliases of the same registered RVT, and a different server carrying the same share-relative path.
- `Needs User Dirty/Revision Check`: test unsaved and unsynchronized project changes, save/sync confirmation, a Standard RVT change immediately before scan, and a Standard RVT change during scan. Dirty or cross-revision data must never become the shared accepted result.
- `Needs User Two-PC Check`: overlap manual/automatic project checks and Standard publishers from two equipped clients. Publication must serialize, the waiting client must reuse the completed result, and the manifest must retain Standard, approved-list, and project-state pointers.
- `Needs User Nested Placement Check`: verify exact, different, and unavailable nested-only fingerprints. Only an exact verified match may be rolled back while Admin Mode is OFF; first-ever placement without precise evidence remains intentionally nonblocking.
- `Needs User Tracking Check`: confirm create/modify/delete/save/sync records, protected deletion marker generations, and catalog-drift detection for work performed by a client without the add-in. An uninstrumented client cannot provide exact actor or transaction time; only the later observed drift can be recorded.
- `Needs User Managed-Folder Check`: disconnect/reconnect the homepage-provided management path and validate test-folder fallback, later migration, offline read-only cache behavior, and latest-revision revalidation immediately before model mutation.
- `Needs User Revit Dialog Check`: verify the hardened scan warning handling against real family Edit/Open warnings, including OK-only and Delete/Cancel-only dialogs, and confirm explicit XLSX export is the only action that creates a workbook.
- `SKIP runtime-not-installed`: Revit 2021 is not installed on this PC. Its shared 2019/2021/2023 assembly, Stage payload, ProgramData payload, static contract, workflow, and representative IE harness are verified.

## 2026-07-22 Family/System Apply And Synchronize Performance Remediation

Status: `Source/Stage/ProgramData Complete / Automated Regression Passed / Needs User Revit Performance Check`

### Checkpoints

- Source checkpoint: `_backups/load-sync-performance-20260722-010118`.
- Deferred-path checkpoint: `_backups/load-sync-performance-network-probe-20260722-0130`.
- ProgramData checkpoint before deployment: `_backups/programdata-before-load-sync-performance-20260722` (`17` files).
- ProgramData checkpoint before the final non-probing build: `_backups/programdata-before-load-sync-network-probe-final-20260722` (`17` files).
- Revit was not launched because automated startup on this PC can trigger the known licence error.

### Confirmed Bottlenecks

- `Family load completion`: the owned Standard RVT remained open while operation history was written, the dashboard/list was rebuilt, the result window was shown, and the user dismissed it. The `finally` block closed the Standard RVT only afterward, so the visible result was not the real terminal point and could be followed by another loading cursor.
- `System Type apply completion`: the same Standard RVT lifetime crossed the result boundary. Selected apply also performed an unconditional full dashboard render before confirmation, and the execution report was serialized twice.
- `Standard path resolution`: a potentially remote `File.Exists(...)` probe ran before checking whether the exact Standard RVT was already open in Revit, adding an avoidable network round trip.
- `Element-history commit construction`: each changed element repeatedly scanned every pending activity to determine ownership and transaction names, producing an avoidable `O(changed elements x activities)` path. Save/Sync also rebuilt project-parameter binding context more than once.
- `Synchronize completion`: after local checkpointing, the Revit completion callback synchronously wrote managed-share history and then enumerated/flushed every pending spool category. A slow or disconnected management share could therefore extend the apparent Synchronize with Central completion time.
- `Residual destination probe`: the first deferred implementation still canonicalized the managed-policy path with `File.Exists/CreateFile` before writing its local envelope. That was enough for an unavailable SMB path to delay the Revit callback even though the actual network write had moved out of it.

### Corrections

- Family load and System Type apply now record explicit timings for registration, Standard snapshot/path/open, execution, pre/post verification, report persistence, Standard RVT close, list refresh, and total ready-for-result time in `%LOCALAPPDATA%\KKY\FamilyBrowser\Diagnostics\dashboard-runtime-YYYYMMDD.log`.
- Both apply flows now finish operation/report persistence, close only the Standard RVT document owned by the command, invalidate/rebuild the affected list, and complete progress before opening the result window. The result window is now the terminal step; dismissing it has no normal follow-up load or refresh.
- Selected System Type apply skips the pre-confirmation full dashboard render. System apply writes its execution report once, after authoritative post-apply verification and tracking are complete.
- Standard document resolution reuses an already-open exact document before probing `File.Exists(...)`; unopened paths retain the availability check and normal open behavior.
- Element-history commit construction builds one activity index containing active/add/delete IDs, per-element activities, and transaction names. Parameter bindings and touched-element capture reuse one `StateCaptureContext` per Save/Sync boundary.
- Element commits are now atomically spooled to the local machine before the Revit callback returns. The callback derives only a non-probing destination token from locally stored configuration text. Managed-path physical canonicalization, availability checks, publication, and pending flush run later on the ThreadPool; duplicate requests for the same root are coalesced, and failures leave validated local spool records for retry. A crash before the worker runs is recovered by the next normal `FlushPending`, which promotes the protected mutable envelope before publication.

### Automated Verification

- Direct Release compilation passed for the three independent host projects. The complete 2019/2021/2023/2025/2027 build and Stage generation then passed with zero compile errors. Its warning set is the existing recovered-source nullable/unused-field, Windows-only API analyzer, and `NETSDK1137` output; no new compile error was introduced.
- Workflow audit passed `1076/1076` checks at `artifacts/family-browser-workflow-audit/20260722-load-sync-performance-final3`. Its executable persistence harness passed `78` checks, including caller-thread provisional destination creation, later atomic canonicalization, offline retention, reconnect flush, and exactly-once immutable history. Static regressions also require result-window terminal ordering, one System apply report write, indexed activity lookup, and one commit capture context.
- Static UI/action and contract checks passed for every host source set: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- The non-Revit 1,000-row performance gate passed at `artifacts/family-browser-ui-audit/20260722-load-sync-performance`: shell `1-2 ms`, cold usable list `415-648 ms`, warm list `280-360 ms`, filter `9-11 ms`, with `150` rows in the DOM.
- Stage and installed add-ins passed verification for Revit 2019/2021/2023/2025/2027. All `17` Stage files match ProgramData by SHA-256 with zero mismatches.
- Installed host DLL SHA-256 values are: 2019/2021/2023 `9161F5239DC736C2A34CBF462B9D2739AE16881268D8BD49DE38C22C17D79A51`; 2025 `01C47483C81575AC96CEA6A8260D15D1A0D6622213265424765D8DCE9002A839`; 2027 `08641BC82BCCFCA4006DFD6C28CBDE0F56E18A7C46EE3232C8F4FBA81A74150C`. Installed version remains `1.0.1`.

### Remaining Runtime Boundary

- `Needs User Family Load Check`: test one and four selected Families against both a local and managed-network Standard RVT. The result window must appear only after list state has changed and dismiss immediately without a second busy cursor.
- `Needs User System Apply Check`: repeat one and four selected System Types. Confirm pre/post verification remains accurate, dependent Families are applied once, the visible result is terminal, and there is no intermediate full-list flash.
- `Needs User Sync Performance Check`: in one guarded workshared project, time a small and a large successful Synchronize before/after this revision. Confirm `%LOCALAPPDATA%\KKY\FamilyBrowser\Diagnostics\element-tracking-performance.log` shows local durable persistence without managed-share wait in the commit callback, and that the shared history appears shortly afterward.
- `Needs User Offline/Two-PC Check`: disconnect the managed share during Sync, reconnect it, and confirm the local spool flushes once without duplicate commits. Then overlap Sync from two equipped clients and confirm both immutable commits remain present. Deferred publication improves UI responsiveness but does not change SMB/server latency or Autodesk's own central synchronization time.

## 2026-07-22 Final Performance And Deployment Review

Status: `Source/Stage/ProgramData Complete / Automated Regression Passed / Needs User Revit Runtime Check`

### Additional Findings And Corrections

- `Pre-progress network wait`: Family load and System Type apply validated the current Standard revision and read snapshot metadata before opening the progress UI. Both paths now start progress before registration/revision validation and close it on validation failure, so a slow Standard share no longer looks like an unresponsive click.
- `Repeated path-identity network probe`: stable path identity repeatedly called `File.Exists/CreateFile`, including Save/Sync identity checks. Successful canonical file or directory identities now use a bounded process-local cache; unavailable paths are deliberately not cached so reconnects and drive remaps can recover.
- `Mapped drive versus UNC Standard reuse`: an already-open Standard RVT was reused only when `Path.GetFullPath` matched exactly. The resolver now falls back to stable canonical identity, preventing a mapped-drive/UNC alias from opening the same Standard document again.
- `ProgramData reinstall`: the deployment script required write access to an administrator-owned `.addin` even when its contents were already identical, then replaced identical protected payloads on every run. It now hashes manifests and payloads, skips unchanged files, removes only genuinely obsolete manifests, and exits successfully when all five installed targets are already current.

### Final Automated Verification

- UI static and contract checks passed for all three host source sets: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- Workflow and durable-tracking audit passed `1087/1087` checks at `artifacts/family-browser-workflow-audit/20260722-015131`; the executable persistence harness also passed.
- The 2,000-row performance/cache gate passed at `artifacts/family-browser-ui-audit/20260722-015035-performance`.
- Targeted IE `WebBrowser` click regression passed `36/36` Korean/English variants across the three independent runtimes at `artifacts/family-browser-ui-audit/20260722-final-targeted-harness`. It covered structured/automatic result windows, Family list and detail preview, System Type detail, and Model Check layout/actions.
- Revit 2019/2021/2023/2025/2027 Release compilation and Stage verification passed with zero errors. Existing recovered-source and Windows-only analyzer warnings remain non-blocking.
- ProgramData verification passed for all five targets. Every staged payload file matches its installed counterpart by SHA-256, and a second non-elevated install run correctly reported all targets as already current.
- Installed version remains `1.0.1`. Final host DLL SHA-256 values are: 2019/2021/2023 `B82643FB8A9DF37FC0FA3613EB9FA6B374D8F1DC685FC33A5AC24E6AC2E4F724`; 2025 `FA5BB208D8447067423C6FF2539B8554587656963D454C086AB122D8EC7C4AB0`; 2027 `5D69BFD6CA2A34A11E2AE582461BFC930B974089788247F5507C1FB00169300C`.

### Runtime Boundary

- Revit was not launched because automated startup on this PC can trigger the known licence error.
- `Needs User Revit Performance Check`: verify one/four Family loads, one/four System Type applies, and guarded Synchronize against local and network Standard paths. The progress UI must appear before remote validation, the result dialog must be terminal, and managed history must publish shortly after the Revit callback without extending it.

## 2026-07-22 Top-Level Model Element History Filtering

Status: `Source/Stage/ProgramData Complete / Automated Regression Passed / Needs User Revit Tracking Check`

### Checkpoints

- Source checkpoint: `_backups\top-level-element-history-filter-20260722-074854`.
- ProgramData checkpoint: `_backups\programdata-before-top-level-history-filter-20260722-080530` (`17` files).
- Revit was not launched because the user performs runtime checks and automated startup on this PC can trigger the known licence error.

### Problem And Correction

- The previous tracking scope accepted almost every positive-ID Revit element with a category. Creating or deleting one user-facing Pipe, Cable Tray, or Conduit could therefore also record logical Piping/Duct systems, centerline graphics, CableTrayRun/ConduitRun containers, and dependent nested Family instances as separate model changes.
- `FamilyBrowserElementTrackingScopePolicy` now defines a stable, exact class/category-ID policy for those Revit support records. Element Type definitions, Families, Materials, Grids, project/shared parameters, and ordinary user-facing instances remain tracked.
- The initial baseline records ignored support IDs in a `HashSet`. `DocumentChanged` resolves each changed element once, drops support IDs before activity, Undo/Redo, touched-ID, and commit bookkeeping, and recognizes later deletions through an O(1) ID lookup without rescanning the document.
- A final commit guard rejects any known support state. The rare late-classification cleanup scans prior activities only when an ignored ID was actually present in earlier touched/ambiguous evidence; normal Pipe/Tray events do not acquire an O(activity count) cleanup path.
- Existing immutable history and its checksum remain untouched. Current Project History, All Project History, preview counts, table rows, and explicit Excel export share one filtered projection. The view reports how many legacy internal-support rows were hidden.
- The excluded scope is intentionally narrow: logical MEP systems, centerline graphics, Cable Tray/Conduit Run containers, and nested Family instances with `SuperComponent` are hidden. Pipe/Duct/Cable Tray/Conduit instances, fittings, insulation, Stair/Railing components, system type definitions, Materials, Grids, and parameter definitions are not broadly suppressed.

### Automated Verification

- Workflow/durable-tracking audit passed `1132/1132` checks at `artifacts\family-browser-workflow-audit\20260722-top-level-history-filter-final2`. It executes the exact scope policy against all `22` excluded BuiltInCategory IDs, auxiliary classes, Korean-name fallbacks, and retained normal/type-definition cases.
- UI static/action and contract checks passed for all three host source sets: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- The 2,000-row performance/cache gate passed at `artifacts\family-browser-ui-audit\20260722-top-level-history-filter-performance-final`; no shell, cold/warm list, filter, cache, or 150-row DOM-window regression was detected.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage verification passed with zero errors. Existing recovered-source and Windows-only analyzer warnings remain non-blocking.
- Elevated ProgramData installation and installed verification passed for all five targets. All `17` Stage files match ProgramData by SHA-256 with zero mismatches. Installed version remains `1.0.1`.
- Final host DLL SHA-256 values: 2019/2021/2023 `3DD0FA104D122CB24742B57B0F1B8E7ED1F7A7CADB4DEF8A60542B8082AB9DE2`; 2025 `FE19BE09B270F90EB95696FE39FEE3A9E30E01A188EBE40A47A3D44B032A0EF5`; 2027 `41AADCAF2F7E12222C2E6C63246F00B9AEBDF7AC38C2109F3EBF99A6497A466B`.

### Runtime Boundary

- `Needs User Revit Tracking Check`: in a guarded RVT, create/modify/delete and Save or Synchronize one Pipe, one Cable Tray, and one Conduit. History must contain the user-facing object and must not contain Centerline, Pipe/Duct System, CableTrayRun, or ConduitRun rows.
- `Needs User Nested History Check`: place and remove a compound Family containing dependent nested instances. Only the top-level Family instance must appear in creation/deletion history; normal independent shared Families must remain visible when they are placed directly.
- `Needs User Undo/Redo Check`: create, undo, redo, delete, and commit the same MEP object. The final change kind and counts must remain exact after auxiliary IDs are removed from the observed event set.
- Existing immutable records can still contain internal rows on disk by design; the history window and its explicit Excel export must hide those rows while reporting the hidden-row count.

### 2026-07-22 Review Follow-up

Status: `Source/Stage/ProgramData Complete / Automated Regression Passed / Needs User Revit Tracking Check`

#### Checkpoints

- Source checkpoint: `_backups\history-filter-review-fixes-20260722-085624`.
- ProgramData checkpoint: `_backups\programdata-before-history-filter-review-fixes-20260722-094334`.
- Revit was not launched because the user performs runtime checks and automated startup on this PC can trigger the known licence error.

#### Corrected Findings

- `All Project History count mismatch`: project-selector totals now run through `FamilyBrowserElementHistoryProjectionPolicy`, the same projection used by detail rows and explicit Excel export. Internal MEP support rows and unresolved same-boundary transient evidence no longer inflate Created/Modified/Deleted totals. Stored immutable commits, their original counters, and checksums remain unchanged.
- `Null live element leakage`: changed-element handling now consults the current-event and session `HashSet` indexes when Revit no longer resolves the live element. Known auxiliary IDs are removed before activity, Undo/Redo, touched-ID, and commit bookkeeping without a document rescan.
- `Unnamed same-boundary rows`: a create/delete sequence with no baseline, current, or last-known metadata remains in immutable evidence but is explicitly marked `UnresolvedTransient`. The shared projection hides it from user-facing counts, tables, and Excel instead of displaying a blank element row.
- `Transition consistency`: add/modify/delete, late-baseline deletion, ambiguous events, and same-boundary create/delete decisions now use `FamilyBrowserElementTrackingTransitionPolicy`. Ambiguous stale activity cannot fabricate a modification, while a real state-signature difference remains visible.
- `Hidden-only empty state`: a history containing only known auxiliary or unresolved transient evidence is now described as retained source evidence with no user-facing change, rather than incorrectly telling the user that tracking never started.
- `Legacy nested-Family boundary`: newly observed dependent nested Family instances are excluded at capture time. Older immutable rows cannot be reliably reclassified because their schema contains no `SuperComponent` or parent marker; the history view now states this limitation instead of guessing or altering checksummed records.
- `Performance boundary`: the normal `DocumentChanged` path still resolves only changed IDs and performs O(1) `HashSet` lookups. The new all-project projection runs only when an administrator opens `전체 프로젝트 이력`; it is not called at startup, Save, or Synchronize.

#### Automated Verification

- Workflow and durable-tracking audit passed `1142/1142` checks at `artifacts\family-browser-workflow-audit\20260722-history-review-fixes-full2`. The executable persistence harness passed `81` behavioral checks, including null live elements, auxiliary filtering, add/modify/delete transitions, ambiguous events, same-boundary transients, immutable-record preservation, and summary/detail projection equality.
- UI static/action and contract checks passed for all three independent host source sets: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host. Nested-family propagation and authoritative System Type apply checks also passed.
- IE `WebBrowser` HTML/click verification completed `232/232` scenarios with zero failures: Revit 2019, 2023, 2025, and 2027 each passed `58/58`. Results and the recovered aggregate are under `artifacts\family-browser-ui-audit\20260722-history-review-fixes-quality-gate2\harness\ui-audit-recovered-summary.json`. Revit 2021 runtime remains `SKIP runtime-not-installed`; its shared assembly and package were still built and verified.
- The final non-Revit quality gate passed at `artifacts\family-browser-ui-audit\20260722-history-review-fixes-quality-gate-final`. The 2,000-row performance/cache gate reported shell `1-2 ms`, cold usable list `439-662 ms`, warm usable list `302-372 ms`, filter `9-12 ms`, and `150` DOM rows, all within acceptance thresholds.
- Revit 2019/2021/2023/2025/2027 Release builds, Stage generation, and staged add-in verification passed with zero errors. Existing recovered-source/analyzer warnings remain non-blocking.

#### ProgramData Deployment

- Elevated installation completed for Revit 2019/2021/2023/2025/2027. `Verify-FamilyBrowserRecovered.ps1 -Installed` passed every target.
- All `17` staged files match ProgramData by SHA-256 with zero missing or mismatched files. Installed version remains `1.0.1`.
- Installed host DLL SHA-256 values: 2019/2021/2023 `438369F6F8025B461880D263A08FDE65E9C83D1728DD27732E2CCEE3EFF5CBD6`; 2025 `3C7F2589E4325960A6691501F105D7676B7977C4065F3810AD4E3AE8E58435DA`; 2027 `81E8335BCD885BE335357BDB634D4E4CC3B626E0F0C12C15468052A91387A539`.

#### Runtime Boundary

- `Needs User Revit Tracking Check`: create/modify/delete and Save or Synchronize one Pipe, Cable Tray, and Conduit. The user-facing object must remain, while logical systems, centerlines, and Run containers must be absent from current/all-project history and exported Excel.
- `Needs User Same-Boundary Check`: create and delete one MEP object before the same successful commit. No unnamed row may appear; the summary may report retained unresolved evidence only when Revit supplied no recoverable metadata.
- `Needs User Nested History Check`: place and remove a compound Family. Newly captured dependent children must be absent while the top-level instance remains. Legacy immutable nested rows may remain because old records cannot prove the parent relationship.

### 2026-07-22 Independent Review After History Filtering

Status: `Review Complete / Source Unchanged / Three Follow-ups Identified`

#### Findings

- `P2 - Electrical logical-system omission`: the scope policy excludes `PipingSystem` and `MechanicalSystem`, but not the third concrete `MEPSystem` subclass, `Autodesk.Revit.DB.Electrical.ElectricalSystem`. Revit 2019/2023/2025/2027 API reflection confirmed the same three concrete subclasses in every installed target. Electrical circuit instances (`OST_ElectricalCircuit`, `-2008037`; internal circuits `-2008152`) can therefore remain visible even though the history UI states that logical MEP systems are excluded. The static scope cases also omit this class and both category IDs.
- `P2 - Same-event auxiliary-to-visible ID transition`: `UpdateCurrentStates(...)` can add an ID to the current-event ignored set while processing deletions, then successfully capture a user-facing live state for that same ID in the added/modified pass without removing the ID from the current-event set. The subsequent `FilterActivityElementIds(...)` and `RemoveIgnoredElementIdsFromSession(...)` calls would discard that state. This requires a rare same-callback ID-role transition, but the tracking ledger should fail closed against it because it can hide a legitimate user-facing change.
- `P3 - All-project history scalability`: `LoadTrackedProjectHistorySummaries(...)` synchronously loads and validates every immutable history JSON, every upload-pending record, and every local checkpoint before showing the project selector. This path is user-invoked and does not affect startup, Save, Synchronize, or `DocumentChanged`, but it can freeze the modeless browser as managed history grows. The current managed fixture has only `8` immutable JSON files (`19,762` bytes), and the automated performance gate does not exercise a large history corpus.

#### Verification

- Independent workflow and durable-tracking audit passed `1142/1142` checks at `artifacts/family-browser-workflow-audit/20260722-independent-review`; its executable persistence harness passed all `81` behavioral checks.
- UI static/action and contract checks passed for all three independent host source sets: `85` generated actions, `265` exact routes, `64` prefix routes, and `11` browser functions per host.
- Installed add-in verification passed for Revit 2019/2021/2023/2025/2027. Revit was not launched because the user performs runtime validation and automated startup on this PC can trigger the known licence error.
- No source, Stage, or ProgramData file was changed during this independent review.

### 2026-07-22 Final 1.0.1 Release Corrections

Status: `Source/Stage/ProgramData/Installer Complete / Automated Regression Passed / Website Published / Needs User Revit Runtime Check`

#### Checkpoints And Corrections

- Source checkpoint: `_backups\history-final-release-fixes-20260722-1025`.
- ProgramData checkpoint: `_backups\programdata-before-familybrowser-v1.0.1-20260722-110119`.
- `Electrical logical-system omission`: tracking scope now excludes `Autodesk.Revit.DB.Electrical.ElectricalSystem`, `OST_ElectricalCircuit` (`-2008037`), internal circuit records (`-2008152`), and the Korean/English electrical-circuit name fallbacks. User-facing electrical model objects remain in scope.
- `Same-event auxiliary-to-visible transition`: a successfully recaptured user-facing state now removes the ID from the current-event and session ignored sets before filtering, so a rare ID-role transition cannot hide a legitimate change.
- `All-project history scalability`: managed immutable history, upload-pending records, and local checkpoints are now read and projected on a background task. The renderer consumes the completed bundle and performs no synchronous history filesystem scan on the modeless UI thread.

#### Automated Verification And Deployment

- Final workflow and durable-tracking audit passed `1149/1149` checks at `artifacts\family-browser-ui-audit\20260722-105836-quality-gate\workflow`.
- The post-format release rerun also passed `1149/1149` checks at `artifacts\family-browser-workflow-audit\20260722-final-release-post-format`.
- Full IE `WebBrowser` HTML/click verification produced `232/232` scenario results with zero failures at `artifacts\family-browser-ui-audit\20260722-102734-quality-gate\harness`. The `32` warnings are expected handled alerts from safe no-selection actions; Revit 2019/2023/2025/2027 each completed `58` scenarios.
- Final non-Revit quality gate passed at `artifacts\family-browser-ui-audit\20260722-105836-quality-gate`, including static/contract routing, nested-family propagation, authoritative System Type apply, five-target Stage verification, and the 2,000-row performance/cache gate.
- Revit 2019/2021/2023/2025/2027 Release builds and Stage verification passed. Elevated ProgramData installation completed and installed verification passed for all five targets. Revit itself was not launched.
- Official installer: `artifacts\family-browser\installers\KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0.1_official-20260722_Setup.exe`; SHA-256 `0183917680C19BD2626C58FD64BA820E1D2B8482DA77018A3B98019C0F637E5A`.
- Mail package: `artifacts\family-browser\mail-packages\20260722_01.zip`; SHA-256 `3F55F1DE05AE1C0ABA2E3B49F0A698663FAE03C263F16C9F8FFAF1B29A9CCDA2`. Distribution audit passed at `artifacts\family-browser-distribution-audit\official-20260722`.
- Homepage release repository commit: `b3ba54ad536c163144352186d423d8e4bb71772a`. GitHub Pages deployment `5548774534` completed successfully. The public feed returned `1.0.1`; the public installer returned `3,887,724` bytes with SHA-256 `0183917680C19BD2626C58FD64BA820E1D2B8482DA77018A3B98019C0F637E5A`; and the published homepage/script pair retained the enabled Family Browser feed-to-download binding.
- Public download: `https://update.zerokky.com/Release/family-browser/official/KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v1.0.1_Setup.exe`.

#### Runtime Boundary

- `Needs User Revit Electrical Tracking Check`: create/modify/delete an electrical circuit and ordinary electrical model objects, then Save/Synchronize. Logical circuit support records must be absent while user-facing objects remain visible.
- `Needs User Revit Same-Event Check`: exercise Undo/Redo and create/delete transitions in one commit boundary; no valid visible state may be discarded by a stale ignored-ID entry.
- `Needs User Large-History Check`: open All Project History against the production managed share and confirm the browser stays responsive while the selector is loading and renders the same projected counts when complete.
