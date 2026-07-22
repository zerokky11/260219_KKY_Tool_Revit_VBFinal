# Family Browser Backup Protocol

Date: 2026-06-26

This file is the short, mandatory rule for future Family Browser edits.

## Why This Exists

The old backup habit copied only the files edited in a small patch. That is not enough when the user asks to return to "the point before the problem", because a visible behavior can depend on several files changed across different turns.

From now on, every Family Browser edit must create both:

1. a full source checkpoint for the Family Browser host source, and
2. a small edit-file backup for only the files about to be changed.

Do not claim "restored to that time" unless the restore target is a full source checkpoint or a verified installed DLL baseline.

## Required Backup Before Any Edit

Before editing any Family Browser source file, create:

```powershell
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$label = "short-task-name"
$checkpoint = "artifacts\source-checkpoints\$stamp-$label"
$editBackup = "artifacts\source-file-backups\$(Get-Date -Format yyyyMMdd)\$stamp-$label"

New-Item -ItemType Directory -Force $checkpoint, $editBackup | Out-Null
```

Then copy these full source folders into the checkpoint:

```powershell
Copy-Item "KKY_FamilyBrowser_RevitHost_2019-2023" "$checkpoint\KKY_FamilyBrowser_RevitHost_2019-2023" -Recurse -Force
Copy-Item "KKY_FamilyBrowser_RevitHost_2025" "$checkpoint\KKY_FamilyBrowser_RevitHost_2025" -Recurse -Force
Copy-Item "KKY_FamilyBrowser_RevitHost_2027" "$checkpoint\KKY_FamilyBrowser_RevitHost_2027" -Recurse -Force
Copy-Item "KKY_FamilyBrowser_Compile" "$checkpoint\KKY_FamilyBrowser_Compile" -Recurse -Force
```

Also copy each file that will be edited into the edit backup, preserving enough path information in the filename or folder name to identify the original.

## Required Manifest

Every checkpoint must include a small manifest:

```powershell
@"
stamp=$stamp
label=$label
reason=<why this checkpoint was made>
source=current working source before edit
edited_files=<list files before editing>
restore_rule=restore full checkpoint only when user asks to return to this exact point
"@ | Set-Content "$checkpoint\CHECKPOINT_MANIFEST.txt" -Encoding UTF8
```

## Restore Rule

When the user says:

- "문제 있기 전으로 되돌려"
- "어제 그 시점으로 되돌려"
- "그때 코드로 되돌려"

Do not use a partial `source-file-backups` folder as if it were the whole source state.

Use this priority:

1. Exact full checkpoint under `artifacts\source-checkpoints\...`
2. Installed ProgramData DLL/PDB baseline decompile, if that exact runtime was accepted
3. Installer/stage artifact from the accepted test point
4. Partial file backups only for the specific files they contain, with a clear warning that this is not a full rollback

## Installer Rule

Do not create an installer immediately after a rollback unless these pass:

```powershell
.\KKY_FamilyBrowser_Compile\Build-FamilyBrowserRecovered.ps1
.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1
```

For recovery-source restores or large Dashboard edits, use a forced rebuild if possible. A normal incremental build is not enough proof.

## Current Lesson

`artifacts\source-file-backups\20260625\20260625-112646-preview-display-fallback` contains only three `FamilyBrowserDashboardHtmlForm.cs` files. It is not a complete source snapshot. Restoring only those files must never be described as restoring the whole add-in to that time.
