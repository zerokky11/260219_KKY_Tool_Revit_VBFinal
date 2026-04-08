param(
    [Parameter(Mandatory = $false)]
    [string]$Version,

    [Parameter(Mandatory = $false)]
    [ValidateSet('release', 'test')]
    [string]$Mode = 'release',

    [Parameter(Mandatory = $false)]
    [string[]]$Attachments,

    [Parameter(Mandatory = $false)]
    [string]$Subject,

    [Parameter(Mandatory = $false)]
    [string]$Body,

    [Parameter(Mandatory = $false)]
    [string]$PasswordPlain,

    [Parameter(Mandatory = $false)]
    [string]$PasswordEnvVar = 'KKY_NATE_SMTP_PASSWORD'
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$releaseDir = Join-Path $repoRoot 'Sever\Release'
$testRoot = Join-Path $repoRoot 'artifacts\test-package'

$smtpHost = 'smtp.mail.nate.com'
$smtpPort = 465
$sender = 'kkykiki89@nate.com'
$recipient = 'kcim03@samoo.co.kr'

function Resolve-AttachmentPaths {
    param(
        [string]$ResolvedVersion,
        [string]$ResolvedMode,
        [string[]]$ExplicitAttachments
    )

    if ($ExplicitAttachments -and $ExplicitAttachments.Count -gt 0) {
        $resolved = @()
        foreach ($path in $ExplicitAttachments) {
            $candidate = if ([System.IO.Path]::IsPathRooted($path)) {
                $path
            } else {
                Join-Path $repoRoot $path
            }

            if (-not (Test-Path -LiteralPath $candidate)) {
                throw "Attachment not found: $candidate"
            }

            $resolved += (Resolve-Path -LiteralPath $candidate).Path
        }
        return $resolved
    }

    if (-not $ResolvedVersion) {
        throw 'Version is required when Attachments are not specified.'
    }

    if ($ResolvedMode -eq 'release') {
        $baseDir = $releaseDir
    } else {
        $baseDir = Join-Path $testRoot $ResolvedVersion
        if (-not (Test-Path -LiteralPath $baseDir)) {
            $baseDir = Join-Path $testRoot 'test'
        }
    }

    $exeName = "KKY_Tool_Revit(2019,21,23,25)_v$ResolvedVersion.exe"
    $zipName = "KKY_Tool_Revit(2019,21,23,25)_v$ResolvedVersion.zip"
    $exePath = Join-Path $baseDir $exeName
    $zipPath = Join-Path $baseDir $zipName

    $found = @()
    foreach ($candidate in @($exePath, $zipPath)) {
        if (Test-Path -LiteralPath $candidate) {
            $found += (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    if ($found.Count -eq 0) {
        throw "No attachments found for version '$ResolvedVersion' in '$baseDir'."
    }

    return $found
}

function Get-Password {
    param(
        [string]$PlainText,
        [string]$EnvName
    )

    if ($PlainText) {
        return $PlainText
    }

    foreach ($scope in 'User', 'Process', 'Machine') {
        $fromEnv = [Environment]::GetEnvironmentVariable($EnvName, $scope)
        if ($fromEnv) {
            return $fromEnv
        }
    }

    throw "SMTP password not provided. Set -PasswordPlain or environment variable '$EnvName'."
}

function Convert-ToMimeEncodedWord {
    param([string]$Text)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $base64 = [Convert]::ToBase64String($bytes)
    return "=?UTF-8?B?$base64?="
}

function New-MimeMessageFile {
    param(
        [string]$From,
        [string]$To,
        [string]$MimeSubject,
        [string]$MimeBody,
        [string[]]$MimeAttachments
    )

    $boundary = "----=_KKY_" + [Guid]::NewGuid().ToString('N')
    $tmpDir = Join-Path $env:TEMP 'KKY_Tool_Revit\Mail'
    New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
    $mimePath = Join-Path $tmpDir ("mail-" + [Guid]::NewGuid().ToString('N') + '.eml')

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine("From: $From")
    [void]$builder.AppendLine("To: $To")
    [void]$builder.AppendLine("Subject: $(Convert-ToMimeEncodedWord $MimeSubject)")
    [void]$builder.AppendLine('MIME-Version: 1.0')
    [void]$builder.AppendLine("Content-Type: multipart/mixed; boundary=`"$boundary`"")
    [void]$builder.AppendLine()

    [void]$builder.AppendLine("--$boundary")
    [void]$builder.AppendLine('Content-Type: text/plain; charset=UTF-8')
    [void]$builder.AppendLine('Content-Transfer-Encoding: base64')
    [void]$builder.AppendLine()
    [void]$builder.AppendLine([Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($MimeBody)))
    [void]$builder.AppendLine()

    foreach ($attachmentPath in $MimeAttachments) {
        $fileName = [System.IO.Path]::GetFileName($attachmentPath)
        $encodedName = Convert-ToMimeEncodedWord $fileName
        $bytes = [System.IO.File]::ReadAllBytes($attachmentPath)
        $base64 = [Convert]::ToBase64String($bytes)

        [void]$builder.AppendLine("--$boundary")
        [void]$builder.AppendLine("Content-Type: application/octet-stream; name=`"$encodedName`"")
        [void]$builder.AppendLine('Content-Transfer-Encoding: base64')
        [void]$builder.AppendLine("Content-Disposition: attachment; filename=`"$encodedName`"")
        [void]$builder.AppendLine()

        for ($i = 0; $i -lt $base64.Length; $i += 76) {
            $length = [Math]::Min(76, $base64.Length - $i)
            [void]$builder.AppendLine($base64.Substring($i, $length))
        }

        [void]$builder.AppendLine()
    }

    [void]$builder.AppendLine("--$boundary--")
    [System.IO.File]::WriteAllText($mimePath, $builder.ToString(), [System.Text.Encoding]::ASCII)
    return $mimePath
}

$resolvedAttachments = Resolve-AttachmentPaths -ResolvedVersion $Version -ResolvedMode $Mode -ExplicitAttachments $Attachments
$resolvedPassword = Get-Password -PlainText $PasswordPlain -EnvName $PasswordEnvVar

if (-not $Subject) {
    if ($Version) {
        $Subject = "KKY Tool Revit $Version build files"
    } else {
        $Subject = 'KKY Tool Revit build files'
    }
}

if (-not $Body) {
    if ($Version) {
        $Body = @"
안녕하세요.

KKY Tool Revit $Version 빌드 파일 전달드립니다.

첨부 파일을 확인해 주세요.
"@
    } else {
        $Body = @"
안녕하세요.

KKY Tool Revit 빌드 파일 전달드립니다.

첨부 파일을 확인해 주세요.
"@
    }
}

$curlPath = (Get-Command curl.exe -ErrorAction Stop).Source
$mimeFile = $null

try {
    $mimeFile = New-MimeMessageFile -From $sender -To $recipient -MimeSubject $Subject -MimeBody $Body -MimeAttachments $resolvedAttachments

    $curlArgs = @(
        '--ssl-reqd'
        '--url', "smtps://$smtpHost`:$smtpPort"
        '--user', "${sender}:$resolvedPassword"
        '--mail-from', $sender
        '--mail-rcpt', $recipient
        '--upload-file', $mimeFile
        '--silent'
        '--show-error'
    )

    & $curlPath @curlArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Mail send failed via curl (exit code: $LASTEXITCODE)"
    }

    Write-Host ''
    Write-Host 'Mail sent successfully.'
    Write-Host "From        : $sender"
    Write-Host "To          : $recipient"
    Write-Host "Subject     : $Subject"
    Write-Host 'Attachments :'
    foreach ($attachmentPath in $resolvedAttachments) {
        Write-Host " - $attachmentPath"
    }
}
finally {
    if ($mimeFile -and (Test-Path -LiteralPath $mimeFile)) {
        Remove-Item -LiteralPath $mimeFile -Force -ErrorAction SilentlyContinue
    }
}
