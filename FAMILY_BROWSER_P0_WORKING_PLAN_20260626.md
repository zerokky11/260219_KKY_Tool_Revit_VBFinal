# Family Browser P0 Working Plan - 2026-06-26

This file is the live working handoff for the current P0 stabilization loop.

## Guardrails

- Do not restore from partial file backups as if they were full source checkpoints.
- Before editing Family Browser source, create a full checkpoint and an edit-file backup.
- Do not rewrite `FamilyBrowserDashboardHtmlForm.cs` wholesale.
- Keep project open and browser startup paths lightweight. No project scan, standard RVT scan, family edit, fingerprint build, or 3D snapshot should run on open.
- Do not create an installer just because the solution builds. Build and verification evidence must be recorded first.

## Current Loop

- Automation id: `family-browser-p0-stabilization-loop`
- Start checkpoint: `artifacts/source-checkpoints/20260626-161738-p0-stabilization-loop-start`
- Start edit backup folder: `artifacts/source-file-backups/20260626/20260626-161738-p0-stabilization-loop-start`

## Intended Behavior

1. Shared managed folder is the operational source of truth for standard scans, project checks, fingerprints, thumbnails, policies, and logs.
2. Revit/project/browser open must only load bounded cached metadata and rows. Heavy scans run only from explicit admin actions.
3. FileGuard blocks native Revit commands only for configured target RVT files. If no FileGuard target is configured, non-admin users are not globally blocked.
4. Family/system load lists are driven by the approved standard Excel/JSON list per trade.
5. A family that is loaded/nested inside another family and also exists in the approved standard list is treated as a nested child and is hidden from the load list.
6. Composite-family detail shows nested child family names only, not nested type names.
7. Detached detail belongs only to Family Load and System Type Load tabs. Other workflow tabs must not open the detail window.
8. Workflow tabs must remain scrollable.
9. Korean/English UI text must come from localized strings and must not contain mojibake.
10. Results, prompts, and settings dialogs should use the consistent HTML-style result surface where practical.

## P0 Order

1. Startup/project-open latency and cache-only behavior.
2. Detached detail window, scroll, and browser tab structure.
3. Nested child filtering and nested detail display.
4. FileGuard target-only behavior and immediate policy refresh.
5. Korean/English text corruption and language toggle consistency.
6. Result/settings dialog consistency.

## Verification Needed

- Code search across 2019-2023, 2025, and 2027 host folders.
- Forced or clean build after any Dashboard or recovery-sensitive edit.
- Family Browser build/stage verification before installer creation.
- External Revit test remains required for UI behavior that cannot be proven by compile.
