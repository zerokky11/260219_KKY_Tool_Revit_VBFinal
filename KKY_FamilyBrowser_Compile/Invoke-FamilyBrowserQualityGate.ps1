param(
    [string[]]$Years = @('2019', '2021', '2023', '2025', '2027'),
    [string]$Configuration = 'Release',
    [string]$OutputDir,
    [switch]$SkipBuild,
    [switch]$SkipHarness,
	[switch]$ManagedDataAudit,
	[switch]$FailWhenManagedDataUnavailable,
    [switch]$Install
)

$ErrorActionPreference = 'Stop'
$Years = @($Years | ForEach-Object { @(([string]$_) -split ',') } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
if (-not $OutputDir) {
    $OutputDir = Join-Path $repoRoot ('artifacts\family-browser-ui-audit\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-quality-gate')
}
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$steps = New-Object System.Collections.Generic.List[object]
$failures = New-Object System.Collections.Generic.List[string]

function Invoke-Step {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Script
    )

    $started = Get-Date
    Write-Host "== $Name ==" -ForegroundColor Cyan
    try {
        & $Script
        $steps.Add([pscustomobject]@{
            Name = $Name
            Status = 'OK'
            StartedAt = $started.ToString('o')
            FinishedAt = (Get-Date).ToString('o')
            Details = ''
        }) | Out-Null
    }
    catch {
        $message = $_.Exception.Message
        $steps.Add([pscustomobject]@{
            Name = $Name
            Status = 'FAIL'
            StartedAt = $started.ToString('o')
            FinishedAt = (Get-Date).ToString('o')
            Details = $message
        }) | Out-Null
        $failures.Add("${Name}: $message") | Out-Null
        throw
    }
}

try {
    Invoke-Step -Name 'UI static and contract checks' -Script {
        $global:LASTEXITCODE = 0
        & (Join-Path $scriptRoot 'Test-FamilyBrowserUiStatic.ps1')
        if ($LASTEXITCODE -ne 0) {
            throw "Test-FamilyBrowserUiStatic.ps1 failed with exit code $LASTEXITCODE"
        }
        $global:LASTEXITCODE = 0
        & (Join-Path $scriptRoot 'Test-NestedFamilyDifferencePropagation.ps1')
        if ($LASTEXITCODE -ne 0) {
            throw "Test-NestedFamilyDifferencePropagation.ps1 failed with exit code $LASTEXITCODE"
        }
        $global:LASTEXITCODE = 0
        & (Join-Path $scriptRoot 'Test-SystemTypeAuthoritativeApply.ps1')
        if ($LASTEXITCODE -ne 0) {
            throw "Test-SystemTypeAuthoritativeApply.ps1 failed with exit code $LASTEXITCODE"
        }
        $global:LASTEXITCODE = 0
        & (Join-Path $scriptRoot 'Test-FamilyBrowserUiContract.ps1') -OutputDir (Join-Path $OutputDir 'contract')
        if ($LASTEXITCODE -ne 0) {
            throw "Test-FamilyBrowserUiContract.ps1 failed with exit code $LASTEXITCODE"
        }
    }

    Invoke-Step -Name 'Run workflow lifecycle and durable tracking audit' -Script {
        $global:LASTEXITCODE = 0
        & (Join-Path $scriptRoot 'Test-FamilyBrowserWorkflow.ps1') -OutputDir (Join-Path $OutputDir 'workflow')
        if ($LASTEXITCODE -ne 0) {
            throw "Test-FamilyBrowserWorkflow.ps1 failed with exit code $LASTEXITCODE"
        }
    }

	if ($ManagedDataAudit) {
		Invoke-Step -Name 'Audit homepage managed-folder data references' -Script {
			$auditArgs = @{
				Years = $Years
				OutputDir = (Join-Path $OutputDir 'managed-data')
			}
			if ($FailWhenManagedDataUnavailable) {
				$auditArgs.TreatUnavailableAsFailure = $true
			}
			& (Join-Path $scriptRoot 'Test-FamilyBrowserManagedData.ps1') @auditArgs
		}
	}

    Invoke-Step -Name 'Build and stage Family Browser addins' -Script {
        $buildArgs = @{
            Years = $Years
            Configuration = $Configuration
        }
        if ($SkipBuild) {
            $buildArgs.SkipBuild = $true
        }
        $global:LASTEXITCODE = 0
        & (Join-Path $scriptRoot 'Build-FamilyBrowserRecovered.ps1') @buildArgs
        if ($LASTEXITCODE -ne 0) {
            throw "Build-FamilyBrowserRecovered.ps1 failed with exit code $LASTEXITCODE"
        }
    }

    Invoke-Step -Name 'Verify staged addins' -Script {
        $global:LASTEXITCODE = 0
        & (Join-Path $scriptRoot 'Verify-FamilyBrowserRecovered.ps1') -Years $Years
        if ($LASTEXITCODE -ne 0) {
            throw "Verify-FamilyBrowserRecovered.ps1 failed with exit code $LASTEXITCODE"
        }
    }

    Invoke-Step -Name 'Run 2,000-row performance and cache gate' -Script {
        $global:LASTEXITCODE = 0
        & (Join-Path $scriptRoot 'Test-FamilyBrowserPerformance.ps1') -Years $Years -OutputDir (Join-Path $OutputDir 'performance')
        if ($LASTEXITCODE -ne 0) {
            throw "Test-FamilyBrowserPerformance.ps1 failed with exit code $LASTEXITCODE"
        }
    }

    if (-not $SkipHarness) {
        Invoke-Step -Name 'Run HTML and click simulation harness' -Script {
            $global:LASTEXITCODE = 0
            & (Join-Path $scriptRoot 'Invoke-FamilyBrowserUiAuditHarness.ps1') -Years $Years -OutputDir (Join-Path $OutputDir 'harness')
            if ($LASTEXITCODE -ne 0) {
                throw "Invoke-FamilyBrowserUiAuditHarness.ps1 failed with exit code $LASTEXITCODE"
            }
        }
    }

    if ($Install) {
        Invoke-Step -Name 'Install addins to ProgramData' -Script {
            $global:LASTEXITCODE = 0
            & (Join-Path $scriptRoot 'Install-FamilyBrowserRecovered.ps1') -Years $Years
            if ($LASTEXITCODE -ne 0) {
                throw "Install-FamilyBrowserRecovered.ps1 failed with exit code $LASTEXITCODE"
            }
        }

        Invoke-Step -Name 'Verify installed addins' -Script {
            $global:LASTEXITCODE = 0
            & (Join-Path $scriptRoot 'Verify-FamilyBrowserRecovered.ps1') -Installed -Years $Years
            if ($LASTEXITCODE -ne 0) {
                throw "Installed Verify-FamilyBrowserRecovered.ps1 failed with exit code $LASTEXITCODE"
            }
        }
    }
}
finally {
    $stepArray = @($steps.ToArray())
    $failureArray = @($failures.ToArray())
    $summary = [ordered]@{
        generatedAt = (Get-Date).ToString('o')
        repoRoot = $repoRoot
        years = @($Years)
        configuration = $Configuration
        install = $Install.IsPresent
		managedDataAudit = $ManagedDataAudit.IsPresent
        outputDir = $OutputDir
        steps = $stepArray
        failures = $failureArray
    }
    $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDir 'quality-gate-summary.json') -Encoding UTF8

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# Family Browser Quality Gate Summary') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("- Generated: $($summary.generatedAt)") | Out-Null
    $lines.Add("- Years: $($Years -join ', ')") | Out-Null
    $lines.Add("- Configuration: $Configuration") | Out-Null
    $lines.Add("- Install: $([bool]$Install)") | Out-Null
	$lines.Add("- Managed data audit: $([bool]$ManagedDataAudit)") | Out-Null
    $lines.Add("- Output: $OutputDir") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('| Step | Status | Details |') | Out-Null
    $lines.Add('|---|---:|---|') | Out-Null
    foreach ($step in $steps) {
        $detail = ([string]$step.Details).Replace('|', '/')
        $lines.Add("| $($step.Name) | $($step.Status) | $detail |") | Out-Null
    }
    if ($failures.Count -gt 0) {
        $lines.Add('') | Out-Null
        $lines.Add('## Failures') | Out-Null
        foreach ($failure in $failures) {
            $lines.Add("- $failure") | Out-Null
        }
    }
    $lines | Set-Content -LiteralPath (Join-Path $OutputDir 'quality-gate-summary.md') -Encoding UTF8
}

if ($failures.Count -gt 0) {
    throw "Family Browser quality gate failed. See $OutputDir"
}

Write-Host "Family Browser quality gate passed: $OutputDir" -ForegroundColor Green
