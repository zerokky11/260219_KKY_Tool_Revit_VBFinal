param(
    [string[]]$Years = @('2019', '2021', '2023', '2025', '2027'),
    [string]$StageRoot,
    [string]$OutputDir,
    [int]$HarnessTimeoutSeconds = 120,
    [int]$StartupShellTargetMs = 500,
    [int]$WarmUsableTargetMs = 1500,
    [int]$ColdUsableTargetMs = 3000,
    [int]$FilterTargetMs = 150,
    [int]$CacheTargetMs = 1500
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Resolve-Path (Join-Path $scriptRoot '..')
if (-not $StageRoot) {
    $StageRoot = Join-Path $repoRoot 'artifacts\family-browser\stage'
}
if (-not $OutputDir) {
    $OutputDir = Join-Path $repoRoot ('artifacts\family-browser-ui-audit\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-performance')
}
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$harnessProject = Join-Path $repoRoot 'KKY_FamilyBrowser_Automation\KKY_FamilyBrowser_UiAuditHarness\KKY_FamilyBrowser_UiAuditHarness.csproj'
& dotnet build $harnessProject -c Release -v:minimal
if ($LASTEXITCODE -ne 0) {
    throw 'UI audit harness build failed.'
}

function Get-HarnessTarget([string]$Year) {
    if (@('2019', '2021', '2023') -contains $Year) { return 'net48' }
    if ($Year -eq '2025') { return 'net8.0-windows' }
    if ($Year -eq '2027') { return 'net10.0-windows' }
    throw "Unsupported Revit year: $Year"
}

function Get-HostAssembly([string]$Year) {
    $folder = Join-Path $StageRoot "Rvt$Year\KKY_FamilyBrowser"
    if (@('2019', '2021', '2023') -contains $Year) { return Join-Path $folder 'KKY_FamilyBrowser_RevitHost.dll' }
    if ($Year -eq '2025') { return Join-Path $folder 'KKY_FamilyBrowser_RevitHost_2025.dll' }
    if ($Year -eq '2027') { return Join-Path $folder 'KKY_FamilyBrowser_RevitHost_2027.dll' }
}

function Add-DependencyDir([System.Collections.Generic.List[string]]$List, [string]$Path) {
    if (-not [string]::IsNullOrWhiteSpace($Path) -and (Test-Path -LiteralPath $Path)) {
        $List.Add($Path) | Out-Null
    }
}

function Get-DependencyDirs([string]$Year, [string]$HostAssembly) {
    $dirs = New-Object System.Collections.Generic.List[string]
    $revitDir = "C:\Program Files\Autodesk\Revit $Year"
    Add-DependencyDir $dirs $revitDir
    Add-DependencyDir $dirs (Join-Path $revitDir 'ko-KR')
    Add-DependencyDir $dirs (Join-Path $revitDir 'en-US')
    Add-DependencyDir $dirs "C:\Program Files\Common Files\Autodesk Shared\RealDWG Shared $Year"
    $componentRoot = "C:\Program Files\Common Files\Autodesk Shared\Components\$Year"
    if (Test-Path -LiteralPath $componentRoot) {
        Add-DependencyDir $dirs $componentRoot
        Get-ChildItem -LiteralPath $componentRoot -Directory -Recurse -ErrorAction SilentlyContinue |
            ForEach-Object { Add-DependencyDir $dirs $_.FullName }
    }
    Add-DependencyDir $dirs (Join-Path $repoRoot "Compile\${Year}addin")
    Add-DependencyDir $dirs (Split-Path -Parent $HostAssembly)
    return $dirs.ToArray()
}

function Join-ProcessArguments([string[]]$Arguments) {
    $quoted = New-Object System.Collections.Generic.List[string]
    foreach ($arg in $Arguments) {
        $quoted.Add('"' + ([string]$arg).Replace('"', '\"') + '"') | Out-Null
    }
    return ($quoted -join ' ')
}

function Invoke-HarnessProcess([string]$Executable, [string[]]$Arguments, [int]$TimeoutSeconds, [string]$StdoutPath, [string]$StderrPath) {
    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = $Executable
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    if ($null -ne $psi.ArgumentList) {
        foreach ($arg in $Arguments) { [void]$psi.ArgumentList.Add($arg) }
    }
    else {
        $psi.Arguments = Join-ProcessArguments $Arguments
    }
    $process = [System.Diagnostics.Process]::Start($psi)
    $timedOut = -not $process.WaitForExit($TimeoutSeconds * 1000)
    if ($timedOut) {
        try { $process.Kill($true) } catch { try { $process.Kill() } catch {} }
    }
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    Set-Content -LiteralPath $StdoutPath -Value $stdout -Encoding UTF8
    Set-Content -LiteralPath $StderrPath -Value $stderr -Encoding UTF8
    return [pscustomobject]@{
        ExitCode = $(if ($timedOut) { 124 } else { $process.ExitCode })
        TimedOut = $timedOut
    }
}

$results = New-Object System.Collections.Generic.List[object]
$failures = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]

foreach ($year in $Years) {
    $hostAssembly = Get-HostAssembly $year
    if (-not (Test-Path -LiteralPath $hostAssembly)) {
        $failures.Add("Rvt$year staged host assembly missing: $hostAssembly") | Out-Null
        continue
    }
    $revitDir = "C:\Program Files\Autodesk\Revit $year"
    if (-not (Test-Path -LiteralPath $revitDir)) {
        $results.Add([pscustomobject]@{ Year = $year; Tab = 'runtime-not-installed'; Status = 'SKIP'; Detail = 'SKIP runtime-not-installed' }) | Out-Null
        continue
    }

    $target = Get-HarnessTarget $year
    $harnessExe = Join-Path $repoRoot "KKY_FamilyBrowser_Automation\KKY_FamilyBrowser_UiAuditHarness\bin\Release\$target\KKY_FamilyBrowser_UiAuditHarness.exe"
    if (-not (Test-Path -LiteralPath $harnessExe)) {
        $failures.Add("Rvt$year harness executable missing: $harnessExe") | Out-Null
        continue
    }
    $dependencyDirs = @(Get-DependencyDirs $year $hostAssembly)
    $dependencyArg = $dependencyDirs -join ';'
    $oldPath = $env:PATH
    $env:PATH = ($dependencyDirs -join [System.IO.Path]::PathSeparator) + [System.IO.Path]::PathSeparator + $oldPath
    try {
        foreach ($tab in @('families', 'systems')) {
            $scenarioDir = Join-Path $OutputDir "Rvt$year-$tab-1000"
            New-Item -ItemType Directory -Force -Path $scenarioDir | Out-Null
            $jsonOut = Join-Path $scenarioDir 'result.json'
            $htmlOut = Join-Path $scenarioDir 'dashboard.html'
            $stdout = Join-Path $scenarioDir 'stdout.txt'
            $stderr = Join-Path $scenarioDir 'stderr.txt'
            $args = @(
                '--assembly', $hostAssembly,
                '--dependencyDir', $dependencyArg,
                '--scenario', "performance-$tab-1000",
                '--workspaceRoot', $repoRoot,
                '--activeTab', $tab,
                '--languageCode', 'ko',
                '--adminMode', 'true',
                '--standardRvtRegistered', 'true',
                '--standardListRegistered', 'true',
                '--includeRows', 'true',
                '--includeRequests', 'false',
                '--includeUnregistered', 'false',
                '--syntheticFamilyCount', '1000',
                '--syntheticSystemCount', '1000',
                '--performanceMode', 'true',
                '--cacheAudit', $(if ($tab -eq 'families') { 'true' } else { 'false' }),
                '--maxClicks', '0',
                '--usableTargetMs', [string]$WarmUsableTargetMs,
                '--filterTargetMs', [string]$FilterTargetMs,
                '--cacheTargetMs', [string]$CacheTargetMs,
                '--width', '1600',
                '--height', '900',
                '--jsonOut', $jsonOut,
                '--htmlOut', $htmlOut
            )
            $processResult = Invoke-HarnessProcess $harnessExe $args $HarnessTimeoutSeconds $stdout $stderr
            if ($processResult.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $jsonOut)) {
                $failures.Add("Rvt$year/$tab performance harness failed with exit code $($processResult.ExitCode). See $scenarioDir") | Out-Null
                continue
            }
            $result = Get-Content -LiteralPath $jsonOut -Raw | ConvertFrom-Json
            $warmUsable = [long]$result.htmlRenderMilliseconds + [long]$result.documentLoadMilliseconds + [long]$result.dashboardReadyMilliseconds
            $coldUsable = [long]$result.coldHtmlRenderMilliseconds + [long]$result.documentLoadMilliseconds + [long]$result.dashboardReadyMilliseconds
            $startupShell = [long]$result.startupShellRenderMilliseconds
            if ($startupShell -gt $StartupShellTargetMs) {
                $failures.Add("Rvt$year/$tab startup shell time ${startupShell}ms exceeded ${StartupShellTargetMs}ms.") | Out-Null
            }
            $status = 'OK'
            if ($coldUsable -gt $ColdUsableTargetMs) {
                $warnings.Add("Rvt$year/$tab first full-list cold goal missed: ${coldUsable}ms exceeded ${ColdUsableTargetMs}ms. Startup shell remained visible during this work.") | Out-Null
                $status = 'WARN'
            }
            $results.Add([pscustomobject]@{
                Year = $year
                Tab = $tab
                Status = $status
                Rows = [int]$result.dataRowCount
                DomRows = [int]$result.domRowCount
                VisibleRows = [int]$result.visibleRowCount
                HtmlLength = [long]$result.htmlLength
                StartupShellMs = $startupShell
                ColdRenderMs = [long]$result.coldHtmlRenderMilliseconds
                WarmRenderMs = [long]$result.htmlRenderMilliseconds
                DocumentMs = [long]$result.documentLoadMilliseconds
                ReadyMs = [long]$result.dashboardReadyMilliseconds
                WarmUsableMs = $warmUsable
                ColdUsableMs = $coldUsable
                FilterMs = [long]$result.filterMilliseconds
                CacheColdMs = [long]$result.cacheColdLoadMilliseconds
                CacheWarmMs = [long]$result.cacheWarmLoadMilliseconds
                CacheOfflineMs = [long]$result.cacheOfflineLoadMilliseconds
                ResultPath = $jsonOut
            }) | Out-Null
        }
    }
    finally {
        $env:PATH = $oldPath
    }
}

$summary = [ordered]@{
    generatedAt = (Get-Date).ToString('o')
    targets = [ordered]@{
        startupShellMs = $StartupShellTargetMs
        warmUsableMs = $WarmUsableTargetMs
        coldUsableMs = $ColdUsableTargetMs
        filterMs = $FilterTargetMs
        cacheMs = $CacheTargetMs
    }
    results = @($results.ToArray())
    warnings = @($warnings.ToArray())
    failures = @($failures.ToArray())
}
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDir 'performance-summary.json') -Encoding UTF8

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('# Family Browser Performance Summary') | Out-Null
$lines.Add('') | Out-Null
$lines.Add("- Generated: $($summary.generatedAt)") | Out-Null
$lines.Add("- Targets: shell <= ${StartupShellTargetMs}ms, warm <= ${WarmUsableTargetMs}ms, first full-list cold goal <= ${ColdUsableTargetMs}ms, filter <= ${FilterTargetMs}ms, cache <= ${CacheTargetMs}ms") | Out-Null
$lines.Add('') | Out-Null
$lines.Add('| Year | Tab | Status | Total | DOM | Visible | Shell | Cold usable | Warm usable | Filter | Cache cold/warm/offline |') | Out-Null
$lines.Add('|---:|---|---:|---:|---:|---:|---:|---:|---:|---:|---|') | Out-Null
foreach ($result in $results) {
    if ($result.Status -eq 'SKIP') {
        $lines.Add("| $($result.Year) | $($result.Tab) | SKIP | - | - | - | - | - | - | - | $($result.Detail) |") | Out-Null
    }
    else {
        $lines.Add("| $($result.Year) | $($result.Tab) | $($result.Status) | $($result.Rows) | $($result.DomRows) | $($result.VisibleRows) | $($result.StartupShellMs)ms | $($result.ColdUsableMs)ms | $($result.WarmUsableMs)ms | $($result.FilterMs)ms | $($result.CacheColdMs)/$($result.CacheWarmMs)/$($result.CacheOfflineMs)ms |") | Out-Null
    }
}
if ($warnings.Count -gt 0) {
    $lines.Add('') | Out-Null
    $lines.Add('## Warnings') | Out-Null
    foreach ($warning in $warnings) { $lines.Add("- $warning") | Out-Null }
}
if ($failures.Count -gt 0) {
    $lines.Add('') | Out-Null
    $lines.Add('## Failures') | Out-Null
    foreach ($failure in $failures) { $lines.Add("- $failure") | Out-Null }
}
$lines | Set-Content -LiteralPath (Join-Path $OutputDir 'performance-summary.md') -Encoding UTF8

if ($failures.Count -gt 0) {
    throw "Family Browser performance gate failed. See $OutputDir"
}
Write-Host "Family Browser performance gate passed: $OutputDir" -ForegroundColor Green
