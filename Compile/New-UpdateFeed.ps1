param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $true)]
    [string]$DownloadUrl,

    [string]$PublishedAt = (Get-Date -Format 'yyyy-MM-dd'),
    [string]$Notes = '',
    [string]$OutputPath = 'latest.json'
)

$payload = [ordered]@{
    version     = $Version
    url         = $DownloadUrl
    publishedAt = $PublishedAt
    notes       = $Notes
}

$json = $payload | ConvertTo-Json -Depth 3
$fullOutputPath = $OutputPath
if (-not [System.IO.Path]::IsPathRooted($fullOutputPath)) {
    $fullOutputPath = Join-Path (Get-Location) $fullOutputPath
}

$outputDir = Split-Path -Parent $fullOutputPath
if ([string]::IsNullOrWhiteSpace($outputDir)) {
    $outputDir = Get-Location
    $fullOutputPath = Join-Path $outputDir (Split-Path -Leaf $OutputPath)
}

if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

[System.IO.File]::WriteAllText($fullOutputPath, $json, [System.Text.Encoding]::UTF8)
Write-Host "Generated update feed:" $fullOutputPath
