# KKY Family Browser Recovery Smoke Test

This checklist is for the recovered decompiled Family Browser hosts.

## 1. Build and stage

```powershell
.\KKY_FamilyBrowser_Compile\Build-FamilyBrowserRecovered.ps1
.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1
```

Expected:
- Revit 2019, 2021, 2023, 2025, and 2027 hosts build with 0 errors.
- Stage verification reports all years as `OK`.
- Staged addin folders do not contain `RevitAPI.dll` or `RevitAPIUI.dll`.

## 2. Install for local Revit test

Close all Revit instances first.

```powershell
.\KKY_FamilyBrowser_Compile\Install-FamilyBrowserRecovered.ps1
.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1 -Installed
```

Expected:
- `C:\ProgramData\Autodesk\Revit\Addins\<year>\KKY_FamilyBrowser` contains the recovered host DLL/PDB/deps files.
- The year root contains the matching `KKY_FamilyBrowser_RevitHost*.addin`.
- Installed verification reports all years as `OK`.

## 3. Revit runtime smoke

For each installed Revit version available on the PC:

1. Start Revit.
2. Confirm the `KKY Browser` ribbon tab appears.
3. Click `Family Browser`.
4. Confirm the dashboard opens without a white-screen hang.
5. If no standard RVT is registered, confirm the UI asks for registration instead of scanning automatically.
6. Open settings and confirm the shared/homepage-managed folder is the active data source.
7. Switch tabs once: Family Load, System Type Load, Requests, Settings.
8. Confirm no heavy scan starts unless a scan button is explicitly pressed.

## 4. Important recovery limits

- This is a decompiled emergency source baseline, not the original source layout.
- Passing build and addin verification proves that Revit can find the host assembly, not that every dashboard feature is behaviorally correct.
- Runtime bugs found after this point should be fixed in the recovered source and immediately committed.
