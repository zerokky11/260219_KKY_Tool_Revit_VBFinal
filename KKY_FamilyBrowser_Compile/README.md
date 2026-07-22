# KKY Family Browser Compile

This folder contains the recovery build path for the decompiled Family Browser Revit hosts.

## Build and stage

```powershell
.\KKY_FamilyBrowser_Compile\Build-FamilyBrowserRecovered.ps1
```

The script builds the recovered hosts and creates this staged layout:

```text
artifacts\family-browser\stage\Rvt2019
artifacts\family-browser\stage\Rvt2021
artifacts\family-browser\stage\Rvt2023
artifacts\family-browser\stage\Rvt2025
artifacts\family-browser\stage\Rvt2027
```

Only Family Browser host files are staged. `RevitAPI.dll` and `RevitAPIUI.dll` are intentionally excluded from the addin folder.

## Verify staged files

```powershell
.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1
```

## Install to Revit addin folders

```powershell
.\KKY_FamilyBrowser_Compile\Install-FamilyBrowserRecovered.ps1
```

This copies the staged files into:

```text
C:\ProgramData\Autodesk\Revit\Addins\<year>\KKY_FamilyBrowser
```

To build and install in one call:

```powershell
.\KKY_FamilyBrowser_Compile\Install-FamilyBrowserRecovered.ps1 -Build
```

## Verify installed files

```powershell
.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1 -Installed
```

## Build installer

```powershell
.\KKY_FamilyBrowser_Compile\Build-FamilyBrowserInstaller.ps1
```

The installer is created under:

```text
artifacts\family-browser\installers
```

The same command also creates a mail-sized zip package larger than 13 MB under:

```text
artifacts\family-browser\mail-packages
```

The zip contains the installer plus `mail_size_padding_do_not_run.bin`, which is only used to trigger large-file mail attachment handling. Use `-SkipMailPackage` only when a mail package is not wanted.
