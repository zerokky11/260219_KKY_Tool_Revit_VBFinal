param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $true)]
    [string]$PublishedAt,

    [string]$Notes = '',

    [string]$PackageUrl = '',

    [string]$InstallerUrl = '',

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [int]$MaxEntries = 30
)

$ErrorActionPreference = 'Stop'

function Split-ReleaseNotes {
    param([string]$Text)

    $normalized = [string]$Text
    $normalized = $normalized.Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @('세부 변경 사항은 업데이트 내역 페이지에서 확인해 주세요.')
    }

    $result = New-Object System.Collections.Generic.List[string]
    foreach ($part in ($normalized -split "(?:\r?\n|\|)")) {
        if (-not $part) { continue }
        if ($part -match "^\r?\n$") { continue }

        $trimmed = $part.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        $null = $result.Add($trimmed)
    }

    if ($result.Count -eq 0) {
        $null = $result.Add('세부 변경 사항은 업데이트 내역 페이지에서 확인해 주세요.')
    }

    return @($result.ToArray())
}

$notesList = Split-ReleaseNotes -Text $Notes

$entry = [ordered]@{
    version      = $Version
    publishedAt  = $PublishedAt
    title        = "Version $Version"
    notes        = $notesList
    packageUrl   = $PackageUrl
    installerUrl = $InstallerUrl
}

$history = @()
if (Test-Path -LiteralPath $OutputPath) {
    try {
        $raw = Get-Content -Raw -LiteralPath $OutputPath
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            $parsed = $raw | ConvertFrom-Json
            if ($parsed -is [System.Collections.IEnumerable]) {
                $history = @($parsed)
            }
        }
    } catch {
        $history = @()
    }
}

$history = @(
    $entry
    $history | Where-Object { $_.version -ne $Version }
)

if ($history.Count -gt $MaxEntries) {
    $history = $history[0..($MaxEntries - 1)]
}

$json = $history | ConvertTo-Json -Depth 6
$utf8WithBom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutputPath, $json, $utf8WithBom)
