# KKY Family Browser Recovery Baseline Lock

Date: 2026-06-24

This file exists to prevent another accidental rebuild from the wrong recovered source tree.

## Current Rule

Do not blindly trust source edits or installers made on 2026-06-24 until they are checked against the installed ProgramData DLL baseline below.

The 2026-06-24 stage/installer DLLs that were tested on another PC are not the source of truth.
They are treated as failed/uncertain delivery artifacts because visible UI behavior was missing there.

The current repair baseline is the DLL set currently installed on this PC under `C:\ProgramData\Autodesk\Revit\Addins\...\KKY_FamilyBrowser`.
Those DLLs were captured and decompiled here:

- `artifacts\installed-programdata-baseline\20260624-102212`

Installed DLL evidence:

- 2019/2021/2023: `KKY_FamilyBrowser_RevitHost.dll`, Length `2463232`, LastWrite `2026-06-22 16:42:18`, SHA256 `C940CA4D156B96D2B157651DC391BC1CC66B2638BE5FFBB5C41CB4420517865B`
- 2025: `KKY_FamilyBrowser_RevitHost_2025.dll`, Length `2461696`, LastWrite `2026-06-22 16:42:20`, SHA256 `84FBC7356EE3AC3CB1443393CFAE99C67035CE629F324DF3C50E2EC50936F48E`
- 2027: `KKY_FamilyBrowser_RevitHost_2027.dll`, Length `2462208`, LastWrite `2026-06-22 16:42:22`, SHA256 `F470E1DC33AD9B3174337FA14FABD07153C12796F5EB9218C12CF43AEFF3C0DA`

The deleted/deprecated 2026-06-22 decompile folder must not be used as a baseline.

Important: this does not mean future fixes should overwrite the latest working source with yesterday's source.
The yesterday baseline is a validation and recovery reference only. Continue fixes from the latest intended working state, and use the baseline to check what drifted or regressed.

Do not make an installer from the current working tree just because it builds. A successful build is not proof that the UI/function baseline is correct.

## Trusted Baseline Order

1. Runtime truth:
   - `C:\ProgramData\Autodesk\Revit\Addins\2019\KKY_FamilyBrowser\KKY_FamilyBrowser_RevitHost.dll`
   - `C:\ProgramData\Autodesk\Revit\Addins\2021\KKY_FamilyBrowser\KKY_FamilyBrowser_RevitHost.dll`
   - `C:\ProgramData\Autodesk\Revit\Addins\2023\KKY_FamilyBrowser\KKY_FamilyBrowser_RevitHost.dll`
   - `C:\ProgramData\Autodesk\Revit\Addins\2025\KKY_FamilyBrowser\KKY_FamilyBrowser_RevitHost_2025.dll`
   - `C:\ProgramData\Autodesk\Revit\Addins\2027\KKY_FamilyBrowser\KKY_FamilyBrowser_RevitHost_2027.dll`

2. Yesterday recovery reports:
   - `artifacts\RESTORE_99_SOURCE_RECOVERY_REPORT.md`
   - `artifacts\restore-pdb-match-report-after-exact.json`
   - `artifacts\restore-pdb-exact-restored-files.json`
   - `artifacts\restore-pdb-mismatches-after-exact-2019-2023.txt`
   - `artifacts\restore-pdb-mismatches-after-exact-2025.txt`
   - `artifacts\restore-pdb-mismatches-after-exact-2027.txt`

3. Yesterday full source candidates:
   - `artifacts\source-backups\20260623-111840-before-sourcelink-restore`
   - `artifacts\source-backups\20260623-110609-before-1642-dll-source-restore`

4. Partial exact-restore backup:
   - `artifacts\source-backups\20260623-114036-before-pdb-exact-restore`
   - This is not a complete source baseline. It contains only the files backed up before exact PDB restore.

## Critical Finding

The 2026-06-23 PDB exact recovery report says the recovery matched almost all files:

- 2019/2021/2023: 218 / 220 matched
- 2025: 215 / 217 matched
- 2027: 215 / 217 matched

But these files did not match PDB checksum and must not be treated as exact source:

- `FamilyBrowserDashboardHtmlForm.cs`
- `FamilyBrowserErrorHelp.cs`

This is important because `FamilyBrowserDashboardHtmlForm.cs` controls most visible browser UI:

- main tab/group layout
- settings/admin layout
- detail panel and detached detail behavior
- standard Excel import flow
- language display
- startup/loading display
- list/detail rendering

If this file is restored from the wrong source, the add-in can still build but lose large UI behavior.

## Current Working Tree Warning

As of this note, the current working tree `FamilyBrowserDashboardHtmlForm.cs` files are not trusted as yesterday's accepted UI baseline.

Observed current 2019-2023 file:

- Path: `KKY_FamilyBrowser_RevitHost_2019-2023\FamilyBrowserDashboardHtmlForm.cs`
- Size: about `1,383,333` bytes
- Modified: 2026-06-24

Observed yesterday candidates:

- `20260623-111840-before-sourcelink-restore`: dashboard about `1,630,688` bytes
- `20260623-110609-before-1642-dll-source-restore`: dashboard about `1,620,047` bytes

The current file is materially smaller, so it must not be assumed equivalent.

## 2026-06-24 ProgramData Baseline Recheck

The active baseline was re-centered on the DLLs currently installed on this PC, not on the 2026-06-24 stage/installer that failed on another PC.

Current finding:

- ProgramData-installed Dashboard decompile is about `1.61 MB`.
- Current working-tree Dashboard source is about `1.38 MB`.
- Therefore a clean compile from the current source is not enough evidence that the delivered UI matches the installed ProgramData baseline.

Known drift confirmed during recheck:

- ProgramData Dashboard contains the HTML worksheet picker for Standard Excel import.
- Current working source had a different sheet prompt path and must not be treated as fully restored.
- Nested-family filtering in current source still needed tightening so stale `IsNestedLoadableChild` flags do not override the standard-list/nested-name rule.

Backup created before the first corrective edit:

- `artifacts\source-file-backups\20260624\20260624-103105-programdata-baseline-nested-guard`

First corrective edit:

- `FamilyBrowserDashboardHtmlForm.cs` in all three Revit host folders now treats a row as a hidden/nested child only when the restricted nested family-name set contains that family name.
- This intentionally removes the fallback that treated `item.IsNestedLoadableChild` by itself as enough.

## 2026-06-24 Recovery Guardrail Update

The attempted Standard Excel HTML sheet-picker merge was backed out after it damaged string/encoding structure in `FamilyBrowserDashboardHtmlForm.cs`.

Current safe state:

- Broken attempt backed up at `artifacts\source-file-backups\20260624\20260624-104448-broken-sheet-picker-attempt`.
- Dashboard files restored from `artifacts\source-file-backups\20260624\20260624-103105-programdata-baseline-nested-guard`.
- Only the known-good nested-child fallback removal was reapplied.
- `Build-FamilyBrowserRecovered.ps1` and `Verify-FamilyBrowserRecovered.ps1` passed after the restore.

Guardrail:

- Do not rewrite the whole decompiled Dashboard file with PowerShell `Set-Content`.
- If merging ProgramData-only UI behavior such as the HTML sheet picker, do it with a small `apply_patch`, or add a separate source file/class and patch only the call site.
- Keep Korean UI strings inside existing encoded files unless the patch is minimal and verified by compile immediately afterward.

## 2026-06-24 Forced Rebuild Finding

The larger Dashboard candidates from the 2026-06-23 backups are useful as reference material, but they are not safe to copy wholesale.

What happened during the baseline restore check:

- Restoring the larger `FamilyBrowserDashboardHtmlForm.cs` made the file closer in size to the installed ProgramData decompile.
- A normal build initially looked successful because the restored source file timestamp was older than the existing output DLL, so MSBuild skipped recompiling it.
- A forced rebuild with `dotnet build -t:Rebuild --no-incremental` exposed syntax errors caused by mojibake-damaged Korean string literals.
- The buildable source was restored from `artifacts\source-file-backups\20260624\20260624-105553-before-dashboard-errorhelp-baseline-restore`.
- The stage was then regenerated and `Verify-FamilyBrowserRecovered.ps1` passed.

Current repair rule:

- Do not use a normal incremental build as proof after copying recovery sources.
- Always run a forced rebuild after any recovery-source restore or large Dashboard/ErrorHelp edit.
- Treat ProgramData decompile and the larger backup files as behavior references.
- Port missing behavior in small patches or separate helper files into the buildable source.
- Do not replace the entire Dashboard file unless the replacement is proven to compile under forced rebuild.

## Mandatory Recovery Procedure

Before changing any Family Browser source:

1. Create a dated backup folder under:
   - `artifacts\source-file-backups\YYYYMMDD\<short-task-name>\`

2. Copy every file that will be edited into the matching relative folder.

3. For recovery work, first compare against the trusted baseline order above.

4. If touching `FamilyBrowserDashboardHtmlForm.cs` or `FamilyBrowserErrorHelp.cs`, explicitly document:
   - source file chosen
   - why it was chosen
   - file size
   - SHA256
   - which UI behaviors were checked

5. Never overwrite all host versions blindly. Check these three locations separately:
   - `KKY_FamilyBrowser_RevitHost_2019-2023`
   - `KKY_FamilyBrowser_RevitHost_2025`
   - `KKY_FamilyBrowser_RevitHost_2027`

6. Do not create an installer until all checks pass:
   - compile succeeds
   - stage verification succeeds
   - UI baseline markers are checked
   - Dashboard source is confirmed against the chosen baseline

## Installer Guard

Before running `Build-FamilyBrowserInstaller.ps1`, run a baseline check and record the result.

Minimum manual checks:

```powershell
Get-FileHash .\KKY_FamilyBrowser_RevitHost_2019-2023\FamilyBrowserDashboardHtmlForm.cs -Algorithm SHA256
Get-Item .\KKY_FamilyBrowser_RevitHost_2019-2023\FamilyBrowserDashboardHtmlForm.cs | Select-Object FullName, Length, LastWriteTime
Get-Content .\artifacts\restore-pdb-match-report-after-exact.json -Raw
Get-Content .\artifacts\restore-pdb-mismatches-after-exact-2019-2023.txt -Raw
```

If the dashboard file is still from the untrusted 2026-06-24 current tree, stop.

## What To Recheck From Yesterday Baseline

Recheck these behaviors against the installed/runtime DLL or accepted yesterday build before further edits:

- browser tab grouping: user work group vs family management group
- language toggle position and Korean/English text
- standard Excel sheet selection flow
- detached detail panel behavior
- admin mode default behavior
- file guard behavior
- standard RVT management panel layout
- family/system type load list layout
- nested family filtering
- startup/loading screen

## Rule For Future Sessions

## 2026-06-26 Open-Latency Guardrail

Do not reintroduce heavy FileGuard work into project-open or dashboard-constructor paths.

The protected-file open freeze was traced to these risk paths:

- `NotifyActiveDocumentChanged()` must stay lightweight. It may set the active document and current user only.
- Dashboard construction must not call `NotifyAdminModeChanged()` in a way that refreshes the protected updater or traverses the Revit ribbon before the shell is visible.
- Native command `CanExecute` must not evaluate policy when `LastActiveDocument` is null.
- Startup overlay dismissal must not wait for a missed JavaScript ready flag when the dashboard root and document body already exist.

Current corrective edits:

- `App.cs` uses `ViewActivated` only to update active document state.
- `FamilyBrowserNativeCommandGuardService.NotifyAdminModeChanged(bool enabled, bool refreshUiNow = true)` supports `refreshUiNow: false`.
- `FamilyBrowserDashboardHtmlForm` constructor calls `NotifyAdminModeChanged(_adminModeEnabled, refreshUiNow: false)`.
- `DismissStartupOverlayIfDashboardReady()` forces `data-kkyfb-ready=true` after document completion if the flag was missed.

Explicit admin toggles and policy changes may still call `NotifyAdminModeChanged()` with the default `refreshUiNow: true` so the user-requested immediate guard/ribbon behavior remains available.

When the user says "yesterday code" or "installed add-in baseline", do not use the current working tree as proof by itself.

Start from:

1. ProgramData installed DLL/PDB
2. 2026-06-23 recovery reports
3. 2026-06-23 source backup folders

Then compare that evidence against the latest intended working source.

Do not roll the source back to yesterday just because this file mentions yesterday's baseline.
Only restore from yesterday's source when the user explicitly asks for rollback or when a specific file is proven corrupted and the restore target is documented.

Then only apply source edits after making a new dated backup.

## 2026-06-24 Latest Validated Patch Layer

After the recovery baseline, the latest validated source layer is the automated UI/guard recheck pass.

Use this as the current working baseline before new edits:

- Backup before edit: `artifacts\source-file-backups\20260624\20260624-112632-auto-p0-ui-guard-pass`
- Audit section: `artifacts\RECOVERY_CURRENT_BASELINE_AUDIT_20260624.md` / `2026-06-24 Automated UI And Guard Recheck`

Validated behaviors in this layer:

- Startup loading labels use non-compatible WinForms text rendering and no source file contains a literal replacement character.
- Detached selected-item detail is opened from the family/system list header and the embedded detail panel is hidden from the browser layout.
- Nested-family display is family-name based and standard-list restricted when a standard list exists.
- File-specific native command guard remains opt-in; if no file guard is configured/enabled, native Load Family/type edit commands are not blocked by this guard.
- Admin-profile users still default to admin mode ON.
- 2027 now has the same embedded dashboard CSS/JS resource path as 2019/2025.

Do not start from an older 2026-06-22 decompile or from an unverified ProgramData DLL when continuing ordinary fixes from this point.

## 2026-06-26 FileGuard Open-Latency Patch

Use this as the current guard-performance layer before future FileGuard edits.

Backup before edit:

- `artifacts\source-file-backups\20260626\20260626-open-fileguard-latency`

Files changed:

- `KKY_FamilyBrowser_RevitHost_2019-2023\FamilyBrowserNativeCommandGuardService.cs`
- `KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserNativeCommandGuardService.cs`
- `KKY_FamilyBrowser_RevitHost_2027\FamilyBrowserNativeCommandGuardService.cs`

Patch intent:

- Revit add-in startup must not read the shared standard policy just to initialize ribbon availability.
- Policy/admin refresh must not build the full protected family/type index.
- `CanExecute` must use lightweight local-path permission evaluation; exact central-path evaluation is kept for actual command execution/blocking.
- `EnsureProtectedElementIndexForGuard()` must not be called from document-open/active/admin-refresh paths.

Verification:

- `dotnet build` passed for 2019-2023, 2025, and 2027 with warnings only.
- `Install-FamilyBrowserRecovered.ps1 -Build` rebuilt and staged successfully, but ProgramData copy was blocked by Windows access to the installed 2019 DLL path. Treat this as deployment permission/lock, not a compile failure.
