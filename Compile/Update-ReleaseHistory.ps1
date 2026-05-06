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

    $fallbackNote = -join ([char[]]@(
        0xC138, 0xBD80, 0x20, 0xBCC0, 0xACBD, 0x20, 0xC0AC, 0xD56D, 0xC740, 0x20,
        0xC5C5, 0xB370, 0xC774, 0xD2B8, 0x20, 0xB0B4, 0xC5ED, 0x20,
        0xD398, 0xC774, 0xC9C0, 0xC5D0, 0xC11C, 0x20, 0xD655, 0xC778, 0xD574, 0x20,
        0xC8FC, 0xC138, 0xC694, 0x002E
    ))

    $normalized = [string]$Text
    $normalized = $normalized.Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @($fallbackNote)
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
        $null = $result.Add($fallbackNote)
    }

    return @($result.ToArray())
}

$notesList = Split-ReleaseNotes -Text $Notes
$titlePrefix = -join ([char[]]@(0xBC84, 0xC804, 0x20))

$entry = [ordered]@{
    version      = $Version
    publishedAt  = $PublishedAt
    title        = "$titlePrefix$Version"
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
