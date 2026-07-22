# Family Browser Intent Audit - 2026-06-26

This document is the live intent matrix for the current stabilization pass.
It exists to prevent future edits from drifting back to older recovered/decompiled behavior.

## Non-Negotiable Working Rules

- Do not use the old 2026-06-22 decompile baseline as an edit base.
- Before source edits, create a full checkpoint under `artifacts/source-checkpoints`.
- Before source edits, create per-file backups under `artifacts/source-file-backups/<date>`.
- Never claim a partial file backup is a full rollback point.
- No heavy project scan, standard RVT scan, family edit, fingerprint build, or 3D snapshot may run on Revit/project/browser open.
- The shared managed folder is the source of truth for standard scans, project scans, thumbnails, fingerprints, policies, and logs.

## Intended Behavior Matrix

| Area | Intended behavior | Current implementation signal | Risk / action |
| --- | --- | --- | --- |
| Project open | Opening a protected RVT must stay lightweight. FileGuard may identify target state but must not scan project families/types. | `NotifyActiveDocumentChanged` is lightweight. `DocumentChanged` checks policy with `includeCentralPath:false`. | Keep this path cache-only. Do not add scan/enumeration here. |
| Browser startup | Startup can load cached rows only. First visible UI must not wait on live project enumeration. | Cache-only branch exists in `RebuildModelerRowsFromRegisteredStandardSlotsLight`. | Preserve this. Add timing logs only if needed. |
| Workflow scroll | Requests, model check, permission/admin, unregistered lists must scroll. | CSS override exists for workflow panes. | Verify all panes keep `overflow:auto`. |
| Detail panel | Family/System load tabs should auto-open/sync detached detail. Other tabs must not open it. | JS only schedules detail sync for `families`/`systems`; right panel hidden on browser tabs. | Add safe fallback if detached window cannot open. |
| Tab layout | Workflow group should include Families, System Types, Requests. Admin group should include Standard RVT, Model Check, Unregistered Families/Types, Permission/Admin, Debug. | Tool rail has this grouping, but old `tabs` row still renders too. | Remove or hide redundant old tab row to reduce visual drift. |
| Nested child filtering | If a family is loaded/nested inside another family and that family name exists in the approved standard list, it is a nested child and must not appear in the load list. | Deep snapshot collects nested `FamilyInstance` and loaded nested `Family`. List filtering restricts nested names to standard Excel index. | Ensure admin rows also exclude nested children, not only modeler rows. |
| Composite detail | Composite detail must list nested family names only. No nested type names. If filtered nested list is empty, show single family. | `BuildNestedLoadableSummary` returns family names only and filters against standard list. | If old scan data contains type text, normalize display and require re-scan only when data is stale. |
| Standard list Excel | Standard list import must select from workbook sheets, not require typed sheet names. | `StandardListSheetSelectionHtmlForm` exists and standard-list import calls it. | Permission Excel still has typed sheet path; standard list is OK. |
| FileGuard | Native Revit command blocking applies only to configured file-specific targets. If no target matches, non-admin users are not globally blocked. | `CanNativeGuardPermission` allows native guard permission when current doc is not file-guard targeted. | Verify immediate refresh after policy changes. |
| FileGuard UI | FileGuard config should open a list-based HTML window directly, with RVT add / folder add / remove / clear / save. | `FileGuardHtmlConfigurationForm` exists and dashboard config uses it. | Keep old folder-first flow out of active actions. |
| CSV / lookup tables | CSV/lookup table content must be part of deep fingerprint so lookup-table edits are detected. | `BuildCoreResultFromOpenFamilyDocument` includes `lookup-tables=`. | Verify project deep scan uses same path as standard deep scan. |
| Result dialogs | Result dialogs should use the modern HTML-like consistent surface where practical. | Family load/model check have custom dialogs; many legacy `MessageBox.Show`/`TaskDialog.Show` remain. | P1 cleanup; do not block P0 on all legacy dialogs. |
| Language | Korean/English visible strings must follow selected language and must not show mojibake. | Many recovered strings still contain mojibake. | Treat as P1 text cleanup, prioritize high-visibility areas first. |

## Current P0 Focus

1. Preserve cache-only startup and project-open behavior.
2. Restore detached detail reliably for Family/System tabs.
3. Restore workflow scrolling.
4. Make nested-child filtering apply in admin and modeler modes.
5. Keep FileGuard target-only command blocking intact.

## 2026-06-26 Stabilization Pass Notes

- Created full checkpoint: `artifacts/source-checkpoints/20260626-163936-p0-dashboard-filter-detail-scroll`.
- Created edit backup: `artifacts/source-file-backups/20260626/20260626-163936-p0-dashboard-filter-detail-scroll`.
- Patched all three Dashboard variants so nested child rows are excluded from loadable counts and admin/modeler family row lists.
- Confirmed workflow scroll and old tab hiding CSS are present in all three Dashboard variants.
- Confirmed `StandardListSheetSelectionHtmlForm` is used for standard list Excel sheet selection.
- Confirmed `FileGuardHtmlConfigurationForm` is the active file-specific guard configuration dialog.
- Confirmed FileGuard native command blocking is target-only in all three host variants through `CanNativeGuardPermission` and `IsFileGuardTargetedDocument`.
- Confirmed deep loadable fingerprint includes lookup table / CSV-size-table data through `lookup-tables=`.
- Verification passed:
  - `dotnet build .\KKY_FamilyBrowser_RevitHost_2025\KKY_FamilyBrowser_RevitHost_2025.csproj -c Release -v:minimal`
  - `.\KKY_FamilyBrowser_Compile\Build-FamilyBrowserRecovered.ps1`
  - `.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1`

## Known Deferred Items

- Full Korean/English string cleanup is broad because recovered strings contain mojibake in many old UI paths.
- Legacy command TaskDialogs still exist outside Dashboard modeless flows.
- External Revit behavior still requires user-side testing after build because WebBrowser/WinForms layout and Revit document events cannot be fully proven by compile.
