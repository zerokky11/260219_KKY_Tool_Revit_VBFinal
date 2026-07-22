# KKY Family Browser Recovery

Recovered on: 2026-06-10 15:03:14 +09:00

The original untracked KKY_FamilyBrowser_* source folders were empty. These files were reconstructed from the latest installed Family Browser DLL/PDB set under C:\ProgramData\Autodesk\Revit\Addins.

Important:
- RevitHost folders are ILSpy C# decompilations, not the original VB source layout.
- Installed DLL/PDB/addin files are preserved under KKY_FamilyBrowser_Compile\RecoveredInstalledAddins.
- Use this as an emergency recovery baseline, then commit these files so the folder cannot disappear again.

2026-06-10 build recovery notes:
- The first recovered C# output did not compile as-is. The buildable baseline was restored from `_recovery_probe` and then minimally patched.
- Required compile fixes:
  - Revit API hint paths were corrected to `..\Compile\<version>addin`.
  - 2025 target framework was changed to `net8.0-windows`.
  - 2027 target framework was changed to `net10.0-windows`.
  - 2027 dashboard CSS/JS resources were restored from the 2025 recovered resources and embedded in the project file.
  - `RevitElementIdCompat` was changed from a decompiled VB module class to a static class.
  - `LoadableFamilyLoadOptions` was adjusted to Revit API `out` signatures.
  - Several decompiled nested helper classes were changed from `private` to `internal` to satisfy generated closure accessibility.
  - Decompiled VB anonymous type references in family cleanup routines were replaced with explicit internal helper classes.
- Verified Release builds:
  - `dotnet build .\KKY_FamilyBrowser_RevitHost_2019-2023\KKY_FamilyBrowser_RevitHost.csproj -c Release -v:minimal`
  - `dotnet build .\KKY_FamilyBrowser_RevitHost_2025\KKY_FamilyBrowser_RevitHost_2025.csproj -c Release -v:minimal`
  - `dotnet build .\KKY_FamilyBrowser_RevitHost_2027\KKY_FamilyBrowser_RevitHost_2027.csproj -c Release -v:minimal`
- Current status: builds pass, but this remains a decompiled emergency source baseline. Revit runtime behavior still needs smoke testing against the installed addin flow before treating it as equivalent to the lost original source.

2026-06-10 recovery environment notes:
- Added Family Browser recovery scripts under `KKY_FamilyBrowser_Compile`.
- Build and stage all supported hosts:
  - `.\KKY_FamilyBrowser_Compile\Build-FamilyBrowserRecovered.ps1`
- Verify staged addin manifests and deployed file set:
  - `.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1`
- Install staged files to `C:\ProgramData\Autodesk\Revit\Addins\<year>`:
  - `.\KKY_FamilyBrowser_Compile\Install-FamilyBrowserRecovered.ps1`
- Verify installed addins:
  - `.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1 -Installed`
- The staging script intentionally excludes `RevitAPI.dll` and `RevitAPIUI.dll` from addin deployment folders.
- Build a distributable installer:
  - `.\KKY_FamilyBrowser_Compile\Build-FamilyBrowserInstaller.ps1`
