# Build Troubleshooting (MSBuild circular dependency / import issues)

## Quick diagnosis script

Use the script below from a **Developer PowerShell** / **PowerShell** on Windows:

```powershell
./tools/msbuild_diagnose.ps1
```

Optional: target a single solution/project.

```powershell
./tools/msbuild_diagnose.ps1 -TargetPath "KKY_Tool_Revit_2025/KKY_Tool_Revit_2025.vbproj"
```

The script generates:

- `*.preprocessed.xml` (`/pp`) for evaluated project/import expansion.
- `*.build.binlog` (`/bl`) for full diagnostics.

Then open `*.build.binlog` in **MSBuild Structured Log Viewer** and inspect the **Imports** tree.

---

## Most common root causes of circular dependency

Circular dependency errors involving SDK internals (for example `Sdk.targets` and `Microsoft.NET.Sdk.DefaultItems.Shared.targets`) are usually caused by one of these:

1. `Directory.Build.props` / `Directory.Build.targets` / custom `*.props` / `*.targets` file declaring SDK via:
   - `<Project Sdk="...">`
2. Manual SDK target imports such as:
   - `...\Sdks\Microsoft.NET.Sdk\Sdk\Sdk.targets`
   - `Microsoft.NET.Sdk.targets`
   - `Microsoft.NET.Sdk.DefaultItems.Shared.targets`

---

## Fix rules (must follow)

- **Only project files** (`.vbproj` / `.csproj`) may declare SDK with `<Project Sdk="...">`.
- `Directory.Build.*` and all `*.props` / `*.targets` must use plain `<Project>` (no SDK attribute).
- Do **not** manually import SDK targets/props; SDK should be brought in only by the project file's SDK declaration.
