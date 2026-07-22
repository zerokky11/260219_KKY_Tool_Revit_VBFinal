param(
    [string]$ContractPath = (Join-Path $PSScriptRoot 'FamilyBrowserUiAudit.contract.json'),
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$contract = Get-Content -LiteralPath $ContractPath -Raw -Encoding UTF8 | ConvertFrom-Json
$failures = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]
$results = New-Object System.Collections.Generic.List[object]

function Add-Failure([string]$message) {
    $script:failures.Add($message) | Out-Null
}

function Add-Warning([string]$message) {
    $script:warnings.Add($message) | Out-Null
}

function ConvertTo-StringSet($values) {
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($value in @($values)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            $set.Add([string]$value) | Out-Null
        }
    }
    return $set
}

function Get-RegexValues([string]$text, [string]$pattern, [int]$group = 1) {
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($match in [regex]::Matches($text, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
        $value = $match.Groups[$group].Value
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $items.Add($value) | Out-Null
        }
    }
    return $items
}

function Test-ActionRouted([string]$action, $exactRoutes, [string[]]$prefixRoutes, $allowedSubwindowActions) {
    if ($exactRoutes.Contains($action)) {
        return $true
    }
    foreach ($prefix in $prefixRoutes) {
        if ($action.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    if ($allowedSubwindowActions.Contains($action)) {
        return $true
    }
    return $false
}

function Test-FunctionDefined([string]$text, [string]$name) {
    $escaped = [regex]::Escape($name)
    return [regex]::IsMatch($text, "function\s+$escaped\s*\(", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase) -or
        [regex]::IsMatch($text, "window\.$escaped\s*=", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

$knownTabs = ConvertTo-StringSet $contract.knownTabs
foreach ($scenario in @($contract.scenarios)) {
    if (-not $knownTabs.Contains([string]$scenario.activeTab)) {
        Add-Failure "Unknown contract scenario tab '$($scenario.activeTab)' in scenario '$($scenario.name)'"
    }
    if ([string]::IsNullOrWhiteSpace([string]$scenario.name)) {
        Add-Failure 'Contract scenario name is empty.'
    }
}

foreach ($folder in @($contract.hostFolders)) {
    $hostPath = Join-Path $repoRoot "$folder\FamilyBrowserDashboardHtmlForm.cs"
    $auditScenarioPath = Join-Path $repoRoot "$folder\FamilyBrowserDashboardAuditScenario.cs"
    $assetPath = Join-Path $repoRoot "$folder\KKY.FamilyBrowser.DashboardAssets.family-browser-shell.js"
    if (-not (Test-Path -LiteralPath $hostPath)) {
        Add-Failure "Missing dashboard host file: $hostPath"
        continue
    }
    if (-not (Test-Path -LiteralPath $auditScenarioPath)) {
        Add-Failure "Missing audit render seam: $auditScenarioPath"
        continue
    }

    $hostText = Get-Content -LiteralPath $hostPath -Raw -Encoding UTF8
    $folderTexts = New-Object System.Collections.Generic.List[string]
    Get-ChildItem -LiteralPath (Join-Path $repoRoot $folder) -Filter '*.cs' -File |
        ForEach-Object {
            $folderTexts.Add((Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8)) | Out-Null
        }
    $auditScenarioText = Get-Content -LiteralPath $auditScenarioPath -Raw -Encoding UTF8
    $assetText = if (Test-Path -LiteralPath $assetPath) { Get-Content -LiteralPath $assetPath -Raw -Encoding UTF8 } else { '' }
    $combinedText = ($folderTexts -join "`n") + "`n" + $assetText

    $exactRoutes = ConvertTo-StringSet (Get-RegexValues $hostText 'case\s+"([^"]+)"\s*:')
    $prefixRoutes = @(Get-RegexValues $hostText 'StartsWith\("([^"]+)"\s*,\s*StringComparison\.OrdinalIgnoreCase\)')
    $allowedSubwindowActions = ConvertTo-StringSet $contract.allowedSubwindowActions
    $generatedActions = ConvertTo-StringSet (Get-RegexValues $combinedText '(?:about:)?kkyfb:([A-Za-z0-9_\-./%]+)')

    foreach ($required in @($contract.requiredExactActions)) {
        if (-not $exactRoutes.Contains([string]$required)) {
            Add-Failure "Required exact route '$required' missing in $folder"
        }
    }

    foreach ($requiredPrefix in @($contract.requiredRoutePrefixes)) {
        if (-not ($prefixRoutes | Where-Object { [string]::Equals($_, [string]$requiredPrefix, [System.StringComparison]::OrdinalIgnoreCase) })) {
            Add-Failure "Required route prefix '$requiredPrefix' missing in $folder"
        }
    }

    foreach ($action in $generatedActions) {
        if ([string]::IsNullOrWhiteSpace($action)) {
            continue
        }
        if (-not (Test-ActionRouted -action $action -exactRoutes $exactRoutes -prefixRoutes $prefixRoutes -allowedSubwindowActions $allowedSubwindowActions)) {
            Add-Failure "Generated kkyfb action '$action' has no main route/prefix or allowed subwindow route in $folder"
        }
    }

    foreach ($functionName in @($contract.requiredBrowserOnlyFunctions)) {
        if (-not (Test-FunctionDefined -text $combinedText -name ([string]$functionName))) {
            Add-Failure "Required browser-only function '$functionName' missing in $folder"
        }
    }

    foreach ($token in @($contract.requiredSourceTokens)) {
        if ($combinedText.IndexOf([string]$token, [System.StringComparison]::Ordinal) -lt 0) {
            Add-Failure "Required source token '$token' missing in $folder"
        }
    }

    foreach ($token in @($contract.forbiddenSourceTokens)) {
        if ($combinedText.IndexOf([string]$token, [System.StringComparison]::Ordinal) -ge 0) {
            Add-Failure "Forbidden source token '$token' present in $folder"
        }
    }

    $results.Add([pscustomobject]@{
        HostFolder = $folder
        GeneratedActionCount = $generatedActions.Count
        ExactRouteCount = $exactRoutes.Count
        PrefixRouteCount = @($prefixRoutes).Count
        BrowserFunctionCount = @($contract.requiredBrowserOnlyFunctions).Count
    }) | Out-Null
}

if ($OutputDir) {
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
    $failureArray = @($failures.ToArray())
    $warningArray = @($warnings.ToArray())
    $resultArray = @($results.ToArray())
    $report = [ordered]@{
        generatedAt = (Get-Date).ToString('o')
        contractPath = (Resolve-Path -LiteralPath $ContractPath).Path
        failures = $failureArray
        warnings = $warningArray
        results = $resultArray
    }
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDir 'ui-contract-report.json') -Encoding UTF8
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# Family Browser UI Contract Report') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("- Generated: $($report.generatedAt)") | Out-Null
    $lines.Add("- Failures: $($failures.Count)") | Out-Null
    $lines.Add("- Warnings: $($warnings.Count)") | Out-Null
    $lines.Add('') | Out-Null
    foreach ($row in $results) {
        $lines.Add("- $($row.HostFolder): actions $($row.GeneratedActionCount), exact routes $($row.ExactRouteCount), prefixes $($row.PrefixRouteCount)") | Out-Null
    }
    if ($failures.Count -gt 0) {
        $lines.Add('') | Out-Null
        $lines.Add('## Failures') | Out-Null
        foreach ($failure in $failures) {
            $lines.Add("- $failure") | Out-Null
        }
    }
    $lines | Set-Content -LiteralPath (Join-Path $OutputDir 'ui-contract-report.md') -Encoding UTF8
}

if ($warnings.Count -gt 0) {
    Write-Host 'UI contract warnings:'
    foreach ($warning in $warnings) {
        Write-Host "WARN $warning"
    }
}

if ($failures.Count -gt 0) {
    Write-Host 'UI contract failures:'
    foreach ($failure in $failures) {
        Write-Host "FAIL $failure"
    }
    exit 1
}

$results | Format-Table -AutoSize
Write-Host 'UI contract checks passed.'
