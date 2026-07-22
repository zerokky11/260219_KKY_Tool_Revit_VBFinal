param(
    [string[]]$Years = @('2019', '2021', '2023', '2025', '2027'),
    [string]$StageRoot,
    [string]$ContractPath = (Join-Path $PSScriptRoot 'FamilyBrowserUiAudit.contract.json'),
    [string]$OutputDir,
    [string]$ScenarioNamePattern,
    [int]$HarnessTimeoutSeconds = 90,
    [switch]$SkipHarnessBuild
)

$ErrorActionPreference = 'Stop'
$Years = @($Years | ForEach-Object { @(([string]$_) -split ',') } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
if (-not $StageRoot) {
    $StageRoot = Join-Path $repoRoot 'artifacts\family-browser\stage'
}
if (-not $OutputDir) {
    $OutputDir = Join-Path $repoRoot ('artifacts\family-browser-ui-audit\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$contract = Get-Content -LiteralPath $ContractPath -Raw -Encoding UTF8 | ConvertFrom-Json
$dashboardScenarios = @($contract.scenarios)
if (-not [string]::IsNullOrWhiteSpace($ScenarioNamePattern)) {
    $dashboardScenarios = @($dashboardScenarios | Where-Object { [string]$_.name -match $ScenarioNamePattern })
    if ($dashboardScenarios.Count -eq 0) {
        throw "No dashboard audit scenarios matched '$ScenarioNamePattern'."
    }
}
$harnessProject = Join-Path $repoRoot 'KKY_FamilyBrowser_Automation\KKY_FamilyBrowser_UiAuditHarness\KKY_FamilyBrowser_UiAuditHarness.csproj'

if (-not $SkipHarnessBuild) {
    & dotnet build $harnessProject -c Release -v:minimal
    if ($LASTEXITCODE -ne 0) {
        throw 'UI audit harness build failed.'
    }
}

function Get-HarnessTarget([string]$Year) {
    if (@('2019', '2021', '2023') -contains $Year) {
        return 'net48'
    }
    if ($Year -eq '2025') {
        return 'net8.0-windows'
    }
    if ($Year -eq '2027') {
        return 'net10.0-windows'
    }
    throw "Unsupported Revit year: $Year"
}

function Get-HostAssembly([string]$Year) {
    $addinFolder = Join-Path $StageRoot "Rvt$Year\KKY_FamilyBrowser"
    if (@('2019', '2021', '2023') -contains $Year) {
        return Join-Path $addinFolder 'KKY_FamilyBrowser_RevitHost.dll'
    }
    if ($Year -eq '2025') {
        return Join-Path $addinFolder 'KKY_FamilyBrowser_RevitHost_2025.dll'
    }
    if ($Year -eq '2027') {
        return Join-Path $addinFolder 'KKY_FamilyBrowser_RevitHost_2027.dll'
    }
    throw "Unsupported Revit year: $Year"
}

function Get-RevitInstallDir([string]$Year) {
    $candidate = "C:\Program Files\Autodesk\Revit $Year"
    if (Test-Path -LiteralPath $candidate) {
        return $candidate
    }
    return ''
}

function Add-DependencyDir([System.Collections.Generic.List[string]]$List, [string]$Path) {
    if (-not [string]::IsNullOrWhiteSpace($Path) -and (Test-Path -LiteralPath $Path)) {
        $List.Add($Path) | Out-Null
    }
}

function Add-AutodeskSharedDependencyDirs([System.Collections.Generic.List[string]]$List, [string]$Year) {
    Add-DependencyDir $List "C:\Program Files\Common Files\Autodesk Shared\RealDWG Shared $Year"

    $componentRoot = "C:\Program Files\Common Files\Autodesk Shared\Components\$Year"
    if (Test-Path -LiteralPath $componentRoot) {
        Add-DependencyDir $List $componentRoot
        Get-ChildItem -LiteralPath $componentRoot -Directory -Recurse -ErrorAction SilentlyContinue |
            ForEach-Object { Add-DependencyDir $List $_.FullName }
    }
}

function Join-ProcessArguments([string[]]$Arguments) {
    $quoted = New-Object System.Collections.Generic.List[string]
    foreach ($arg in $Arguments) {
        $value = [string]$arg
        $value = $value.Replace('"', '\"')
        $quoted.Add('"' + $value + '"') | Out-Null
    }
    return ($quoted -join ' ')
}

function Get-ScenarioLanguageVariants($Scenario) {
    $languages = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in @(([string]$Scenario.languageCode), 'ko', 'en')) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }
        $code = if ($candidate -ieq 'en') { 'en' } else { 'ko' }
        if (-not $languages.Contains($code)) {
            $languages.Add($code) | Out-Null
        }
    }
    return $languages.ToArray()
}

function Get-ScenarioThemeVariants($Scenario) {
    return @('light')
}

function Invoke-HarnessProcess([string]$Executable, [string[]]$Arguments, [int]$TimeoutSeconds, [string]$StdoutPath, [string]$StderrPath) {
    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = $Executable
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    if ($null -ne $psi.ArgumentList) {
        foreach ($arg in $Arguments) {
            [void]$psi.ArgumentList.Add($arg)
        }
    }
    else {
        $psi.Arguments = Join-ProcessArguments $Arguments
    }

    $process = [System.Diagnostics.Process]::Start($psi)
    $timedOut = -not $process.WaitForExit($TimeoutSeconds * 1000)
    if ($timedOut) {
        try {
            $process.Kill($true)
        }
        catch {
            try { $process.Kill() } catch {}
        }
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

function Get-EdgeExecutable {
    $candidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    )
    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }
    return ''
}

function Invoke-HtmlScreenshot([string]$HtmlPath, [string]$PngPath, [int]$Width = 1180, [int]$Height = 900) {
    $edge = Get-EdgeExecutable
    if ([string]::IsNullOrWhiteSpace($edge) -or -not (Test-Path -LiteralPath $HtmlPath)) {
        return $false
    }
    if ($Width -le 0) { $Width = 1180 }
    if ($Height -le 0) { $Height = 900 }
    $htmlFullPath = (Resolve-Path -LiteralPath $HtmlPath).Path
    $pngFullPath = [System.IO.Path]::GetFullPath($PngPath)
    $captureId = [Guid]::NewGuid().ToString('N')
    $profileDir = Join-Path $env:TEMP ('kkyfb-edge-' + $captureId)
    $temporaryPng = Join-Path $env:TEMP ('kkyfb-edge-' + $captureId + '.png')
    New-Item -ItemType Directory -Force -Path $profileDir | Out-Null
    $uri = ([System.Uri]::new($htmlFullPath)).AbsoluteUri
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--allow-file-access-from-files', '--run-all-compositor-stages-before-draw', '--virtual-time-budget=1500', "--window-size=$Width,$Height", "--user-data-dir=$profileDir", "--screenshot=$temporaryPng", $uri)
    # Edge writes a successful "bytes written" message to stderr. With the
    # quality gate's Stop preference, PowerShell can promote that message to a
    # terminating NativeCommandError before the PNG check below runs.
    $previousErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'SilentlyContinue'
        & $edge @edgeArgs 2>$null | Out-Null
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
    $deadline = [DateTime]::UtcNow.AddSeconds(12)
    while ([DateTime]::UtcNow -lt $deadline -and -not (Test-Path -LiteralPath $temporaryPng)) {
        Start-Sleep -Milliseconds 150
    }
    if (-not (Test-Path -LiteralPath $temporaryPng) -or (Get-Item -LiteralPath $temporaryPng).Length -le 0) {
        return $false
    }
    [System.IO.File]::Copy($temporaryPng, $pngFullPath, $true)
    return (Test-Path -LiteralPath $pngFullPath) -and (Get-Item -LiteralPath $pngFullPath).Length -gt 0
}

function Invoke-DetailHtmlScreenshot([string]$HtmlPath, [string]$PngPath) {
    return Invoke-HtmlScreenshot -HtmlPath $HtmlPath -PngPath $PngPath -Width 1180 -Height 900
}

$results = New-Object System.Collections.Generic.List[object]
$failures = New-Object System.Collections.Generic.List[string]

foreach ($year in $Years) {
    $target = Get-HarnessTarget $year
    $harnessExe = Join-Path $repoRoot "KKY_FamilyBrowser_Automation\KKY_FamilyBrowser_UiAuditHarness\bin\Release\$target\KKY_FamilyBrowser_UiAuditHarness.exe"
    $hostAssembly = Get-HostAssembly $year
    if (-not (Test-Path -LiteralPath $harnessExe)) {
        throw "Missing harness executable: $harnessExe"
    }
    if (-not (Test-Path -LiteralPath $hostAssembly)) {
        throw "Missing staged host assembly: $hostAssembly"
    }

    $revitInstallDir = Get-RevitInstallDir $year
    if (-not $revitInstallDir) {
        $message = "Revit $year install directory was not found; staged addin verification still covers the manifest/package, but UI harness runtime smoke is skipped for this year."
        Write-Warning $message
        $results.Add([pscustomobject]@{
            Year = $year
            Scenario = 'runtime-not-installed'
            Status = 'SKIP'
            ClickableCount = 0
            BrowserClickCount = 0
            HostActionCandidateCount = 0
            FailureCount = 0
            WarningCount = 1
            ResultPath = ''
        }) | Out-Null
        continue
    }

    $dependencyParts = New-Object System.Collections.Generic.List[string]
    Add-DependencyDir $dependencyParts $revitInstallDir
    Add-DependencyDir $dependencyParts (Join-Path $revitInstallDir 'ko-KR')
    Add-DependencyDir $dependencyParts (Join-Path $revitInstallDir 'en-US')
    Add-AutodeskSharedDependencyDirs $dependencyParts $year
    Add-DependencyDir $dependencyParts (Join-Path $repoRoot "Compile\${year}addin")
    Add-DependencyDir $dependencyParts (Split-Path -Parent $hostAssembly)
    $dependencyDir = ($dependencyParts.ToArray() -join ';')
    $oldPath = $env:PATH
    $env:PATH = ($dependencyParts.ToArray() -join [System.IO.Path]::PathSeparator) + [System.IO.Path]::PathSeparator + $oldPath
    try {
        foreach ($languageCode in @('ko', 'en')) {
          foreach ($themeCode in @('light')) {
           foreach ($messageFixture in @('structured', 'auto-result')) {
            $scenarioName = if ($messageFixture -eq 'auto-result') { "auto-result-message-$languageCode-$themeCode" } else { "structured-message-$languageCode-$themeCode" }
            $scenarioDir = Join-Path $OutputDir ("Rvt$year-$scenarioName")
            New-Item -ItemType Directory -Force -Path $scenarioDir | Out-Null
            $jsonOut = Join-Path $scenarioDir 'result.json'
            $htmlOut = Join-Path $scenarioDir 'message.html'
            $args = @(
                '--assembly', $hostAssembly,
                '--dependencyDir', $dependencyDir,
                '--scenario', $scenarioName,
                '--renderMode', 'message',
                '--messageFixture', $messageFixture,
                '--languageCode', $languageCode,
                '--themeCode', $themeCode,
                '--jsonOut', $jsonOut,
                '--htmlOut', $htmlOut
            )
			if ($languageCode -eq 'ko' -and $messageFixture -eq 'structured') {
				$args += @('--revisionAudit', 'true')
			}
            $stdoutPath = Join-Path $scenarioDir 'stdout.txt'
            $stderrPath = Join-Path $scenarioDir 'stderr.txt'
            $run = Invoke-HarnessProcess -Executable $harnessExe -Arguments $args -TimeoutSeconds $HarnessTimeoutSeconds -StdoutPath $stdoutPath -StderrPath $stderrPath
            $exitCode = $run.ExitCode
            if (Test-Path -LiteralPath $jsonOut) {
                $resultJson = Get-Content -LiteralPath $jsonOut -Raw -Encoding UTF8 | ConvertFrom-Json
                $scenarioFailures = New-Object System.Collections.Generic.List[string]
                foreach ($failure in @($resultJson.failures)) {
                    $scenarioFailures.Add([string]$failure) | Out-Null
                }
                if ($run.TimedOut) {
                    $scenarioFailures.Add("harness process timed out after $HarnessTimeoutSeconds seconds.") | Out-Null
                }
                $results.Add([pscustomobject]@{
                    Year = $year
                    Scenario = $scenarioName
                    Status = $(if ($exitCode -eq 0 -and $scenarioFailures.Count -eq 0) { 'OK' } else { 'FAIL' })
                    ClickableCount = $resultJson.clickableCount
                    BrowserClickCount = $resultJson.browserClickCount
                    HostActionCandidateCount = $resultJson.hostActionCandidateCount
                    FailureCount = $scenarioFailures.Count
                    WarningCount = @($resultJson.warnings).Count
                    ResultPath = $jsonOut
                }) | Out-Null
                foreach ($failure in @($scenarioFailures.ToArray())) {
                    $failures.Add("Rvt$year/${scenarioName}: $failure") | Out-Null
                }
            }
            else {
                $results.Add([pscustomobject]@{
                    Year = $year
                    Scenario = $scenarioName
                    Status = 'FAIL'
                    ClickableCount = 0
                    BrowserClickCount = 0
                    HostActionCandidateCount = 0
                    FailureCount = 1
                    WarningCount = 0
                    ResultPath = $jsonOut
                }) | Out-Null
                $failures.Add("Rvt$year/${scenarioName}: harness did not write result JSON.") | Out-Null
            }
           }
          }
        }

        foreach ($scenario in $dashboardScenarios) {
            $baseScenarioName = [string]$scenario.name
            foreach ($languageCode in @(Get-ScenarioLanguageVariants $scenario)) {
              foreach ($themeCode in @(Get-ScenarioThemeVariants $scenario)) {
            $scenarioName = "$baseScenarioName-$languageCode-$themeCode"
            $scenarioDir = Join-Path $OutputDir ("Rvt$year-$scenarioName")
            New-Item -ItemType Directory -Force -Path $scenarioDir | Out-Null
            $jsonOut = Join-Path $scenarioDir 'result.json'
            $htmlOut = Join-Path $scenarioDir 'dashboard.html'
            $dashboardVisualHtmlOut = ''
            $dashboardPngOut = ''
            $detailHtmlOut = ''
            $detailPngOut = ''
            $initialLanguageCode = if ($languageCode -eq 'en') { 'ko' } else { [string]$languageCode }
            $args = @(
                '--assembly', $hostAssembly,
                '--dependencyDir', $dependencyDir,
                '--workspaceRoot', $repoRoot,
                '--scenario', $scenarioName,
                '--activeTab', ([string]$scenario.activeTab),
                '--languageCode', ([string]$languageCode),
                '--initialLanguageCode', $initialLanguageCode,
                '--themeCode', ([string]$themeCode),
                '--adminMode', ([string]$scenario.adminMode),
                '--standardRvtRegistered', ([string]$scenario.standardRvtRegistered),
                '--standardListRegistered', ([string]$scenario.standardListRegistered),
                '--standardRvtChanged', ([string]$scenario.standardRvtChanged),
                '--standardRvtUnavailable', ([string]$scenario.standardRvtUnavailable),
                '--includeRows', ([string]$scenario.includeRows),
                '--includeRequests', ([string]$scenario.includeRequests),
                '--includeUnregistered', ([string]$scenario.includeUnregistered),
                '--includeReadinessWarning', ([string]$scenario.includeReadinessWarning),
                '--userIdentity', ([string]$scenario.userIdentity),
                '--width', ([string]$scenario.width),
                '--height', ([string]$scenario.height),
                '--jsonOut', $jsonOut,
                '--htmlOut', $htmlOut
            )
            if ($null -ne $scenario.PSObject.Properties['projectPath']) {
                $args += @('--projectPath', ([string]$scenario.projectPath))
            }
            if ($null -ne $scenario.PSObject.Properties['centralPath']) {
                $args += @('--centralPath', ([string]$scenario.centralPath))
            }
            if ($null -ne $scenario.PSObject.Properties['browseDisciplineKey']) {
                $args += @('--browseDisciplineKey', ([string]$scenario.browseDisciplineKey))
            }
            if ($null -ne $scenario.PSObject.Properties['policyActiveDisciplineKey']) {
                $args += @('--policyActiveDisciplineKey', ([string]$scenario.policyActiveDisciplineKey))
            }
            if ($null -ne $scenario.PSObject.Properties['includePendingRows']) {
                $args += @('--includePendingRows', ([string]$scenario.includePendingRows))
            }
            if ($null -ne $scenario.PSObject.Properties['managedFolderUnavailable']) {
                $args += @('--managedFolderUnavailable', ([string]$scenario.managedFolderUnavailable))
            }
            if ($null -ne $scenario.PSObject.Properties['managedFolderTestOverride']) {
                $args += @('--managedFolderTestOverride', ([string]$scenario.managedFolderTestOverride))
            }
            if ($null -ne $scenario.PSObject.Properties['homepageManagedFolderAvailable']) {
                $args += @('--homepageManagedFolderAvailable', ([string]$scenario.homepageManagedFolderAvailable))
            }
            if ($null -ne $scenario.PSObject.Properties['projectCatalogBaselineMissing']) {
                $args += @('--projectCatalogBaselineMissing', ([string]$scenario.projectCatalogBaselineMissing))
            }
            if ($null -ne $scenario.PSObject.Properties['projectCatalogChanged']) {
                $args += @('--projectCatalogChanged', ([string]$scenario.projectCatalogChanged))
            }
            if ($null -ne $scenario.PSObject.Properties['projectCatalogUntracked']) {
                $args += @('--projectCatalogUntracked', ([string]$scenario.projectCatalogUntracked))
            }
            if ($null -ne $scenario.PSObject.Properties['trackingPendingCount']) {
                $args += @('--trackingPendingCount', ([string]$scenario.trackingPendingCount))
            }
            if ($null -ne $scenario.PSObject.Properties['compareDetailedSystemTypeComponents']) {
                $args += @('--compareDetailedSystemTypeComponents', ([string]$scenario.compareDetailedSystemTypeComponents))
            }
            if ($languageCode -eq 'ko' -and $baseScenarioName -in @('admin-home-with-data', 'admin-home-managed-folder-unavailable', 'admin-home-test-folder-homepage-available', 'admin-standard-settings-layout', 'viewport-1280-family', 'viewport-1600-home', 'viewport-1920-admin')) {
                $args += @('--screenshotOut', (Join-Path $scenarioDir 'preview.png'))
                $dashboardVisualHtmlOut = Join-Path $scenarioDir 'dashboard-preview.html'
                $dashboardPngOut = Join-Path $scenarioDir 'preview-edge.png'
                $args += @('--visualHtmlOut', $dashboardVisualHtmlOut)
            }
            if ($languageCode -eq 'ko' -and $baseScenarioName -in @('admin-family-detail-with-preview', 'admin-system-with-data')) {
                $detailHtmlOut = Join-Path $scenarioDir 'detail-preview.html'
                $detailPngOut = Join-Path $scenarioDir 'detail-preview.png'
                $args += @('--detailHtmlOut', $detailHtmlOut)
            }

            $stdoutPath = Join-Path $scenarioDir 'stdout.txt'
            $stderrPath = Join-Path $scenarioDir 'stderr.txt'
            $run = Invoke-HarnessProcess -Executable $harnessExe -Arguments $args -TimeoutSeconds $HarnessTimeoutSeconds -StdoutPath $stdoutPath -StderrPath $stderrPath
            $exitCode = $run.ExitCode
            $dashboardScreenshotFailure = ''
            if (-not [string]::IsNullOrWhiteSpace($dashboardPngOut)) {
                $dashboardWidth = if ([int]$scenario.width -gt 0) { [int]$scenario.width } else { 1740 }
                $dashboardHeight = if ([int]$scenario.height -gt 0) { [int]$scenario.height } else { 980 }
                if (-not (Invoke-HtmlScreenshot -HtmlPath $dashboardVisualHtmlOut -PngPath $dashboardPngOut -Width $dashboardWidth -Height $dashboardHeight)) {
                    $dashboardScreenshotFailure = 'standalone dashboard screenshot failed.'
                }
            }
            $detailScreenshotFailure = ''
            if (-not [string]::IsNullOrWhiteSpace($detailHtmlOut)) {
                if (-not (Invoke-DetailHtmlScreenshot -HtmlPath $detailHtmlOut -PngPath $detailPngOut)) {
                    $detailScreenshotFailure = 'standalone detached detail screenshot failed.'
                }
            }
            if (Test-Path -LiteralPath $jsonOut) {
                $resultJson = Get-Content -LiteralPath $jsonOut -Raw -Encoding UTF8 | ConvertFrom-Json
                $scenarioFailures = New-Object System.Collections.Generic.List[string]
                foreach ($failure in @($resultJson.failures)) {
                    $scenarioFailures.Add([string]$failure) | Out-Null
                }
                if ($run.TimedOut) {
                    $scenarioFailures.Add("harness process timed out after $HarnessTimeoutSeconds seconds.") | Out-Null
                }
                if (-not [string]::IsNullOrWhiteSpace($dashboardScreenshotFailure)) {
                    $scenarioFailures.Add($dashboardScreenshotFailure) | Out-Null
                }
                if (-not [string]::IsNullOrWhiteSpace($detailScreenshotFailure)) {
                    $scenarioFailures.Add($detailScreenshotFailure) | Out-Null
                }
                $results.Add([pscustomobject]@{
                    Year = $year
                    Scenario = $scenarioName
                    Status = $(if ($exitCode -eq 0 -and $scenarioFailures.Count -eq 0) { 'OK' } else { 'FAIL' })
                    ClickableCount = $resultJson.clickableCount
                    BrowserClickCount = $resultJson.browserClickCount
                    HostActionCandidateCount = $resultJson.hostActionCandidateCount
                    FailureCount = $scenarioFailures.Count
                    WarningCount = @($resultJson.warnings).Count
                    ResultPath = $jsonOut
                }) | Out-Null
                foreach ($failure in @($scenarioFailures.ToArray())) {
                    $failures.Add("Rvt$year/${scenarioName}: $failure") | Out-Null
                }
            }
            else {
                $results.Add([pscustomobject]@{
                    Year = $year
                    Scenario = $scenarioName
                    Status = 'FAIL'
                    ClickableCount = 0
                    BrowserClickCount = 0
                    HostActionCandidateCount = 0
                    FailureCount = 1
                    WarningCount = 0
                    ResultPath = $jsonOut
                }) | Out-Null
                $failures.Add("Rvt$year/${scenarioName}: harness did not write result JSON.") | Out-Null
            }
              }
            }
        }
    }
    finally {
        $env:PATH = $oldPath
    }
}

$summary = [ordered]@{
    generatedAt = (Get-Date).ToString('o')
    contractPath = (Resolve-Path -LiteralPath $ContractPath).Path
    outputDir = (Resolve-Path -LiteralPath $OutputDir).Path
    results = @($results.ToArray())
    failures = @($failures.ToArray())
}
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDir 'ui-harness-summary.json') -Encoding UTF8

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('# Family Browser UI Harness Summary') | Out-Null
$lines.Add('') | Out-Null
$lines.Add("- Generated: $($summary.generatedAt)") | Out-Null
$lines.Add("- Output: $OutputDir") | Out-Null
$lines.Add("- Failures: $($failures.Count)") | Out-Null
$lines.Add('') | Out-Null
$lines.Add('| Year | Scenario | Status | Clickables | Browser Clicks | Host Actions | Failures | Warnings |') | Out-Null
$lines.Add('|---|---|---:|---:|---:|---:|---:|---:|') | Out-Null
foreach ($row in $results) {
    $lines.Add("| $($row.Year) | $($row.Scenario) | $($row.Status) | $($row.ClickableCount) | $($row.BrowserClickCount) | $($row.HostActionCandidateCount) | $($row.FailureCount) | $($row.WarningCount) |") | Out-Null
}
if ($failures.Count -gt 0) {
    $lines.Add('') | Out-Null
    $lines.Add('## Failures') | Out-Null
    foreach ($failure in $failures) {
        $lines.Add("- $failure") | Out-Null
    }
}
$lines | Set-Content -LiteralPath (Join-Path $OutputDir 'ui-harness-summary.md') -Encoding UTF8

$results | Format-Table -AutoSize
if ($failures.Count -gt 0) {
    throw 'Family Browser UI harness checks failed.'
}

Write-Host "Family Browser UI harness checks passed: $OutputDir" -ForegroundColor Green
