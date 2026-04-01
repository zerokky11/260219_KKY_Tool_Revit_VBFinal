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

function Convert-TextToJsonSafeAscii {
    param([string]$Text)

    if ($null -eq $Text) {
        return ''
    }

    $builder = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int][char]$ch
        if ($code -gt 127) {
            [void]$builder.AppendFormat('\u{0:x4}', $code)
        } else {
            [void]$builder.Append($ch)
        }
    }

    return $builder.ToString()
}

function Split-ReleaseNotes {
    param([string]$Text)

    $normalized = [string]$Text
    $normalized = $normalized.Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @('세부 변경 사항은 내부 기록을 확인해 주세요.')
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
        $null = $result.Add('세부 변경 사항은 내부 기록을 확인해 주세요.')
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
$json = Convert-TextToJsonSafeAscii -Text $json
[System.IO.File]::WriteAllText($OutputPath, $json, [System.Text.Encoding]::UTF8)
