#Requires -Version 5.1
<#
.SYNOPSIS
    Auto-locating launcher for Invoke-MailboxArchive.ps1 (Mailbox Archive Runner for Exchange Online).

.DESCRIPTION
    Finds Invoke-MailboxArchive.ps1 next to this launcher and runs it, passing
    through every argument unchanged. When the script is not present locally it
    is downloaded from the repository's raw URL first, so the launcher works
    from any folder and always runs the current version.

.EXAMPLE
    .\Run-MailboxArchive.ps1

.EXAMPLE
    .\Run-MailboxArchive.ps1 -Mailbox user@contoso.com -RetentionPolicy "Default MRM Policy" -Passes 3 -IntervalMinutes 10

.EXAMPLE
    .\Run-MailboxArchive.ps1 -Mailbox user@contoso.com -Mode Check

.EXAMPLE
    .\Run-MailboxArchive.ps1 -Mailbox user@contoso.com -ClearRetentionHold -Passes 6 -IntervalMinutes 15
#>
param()

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$local = Join-Path $scriptDir 'Invoke-MailboxArchive.ps1'

if (Test-Path -LiteralPath $local) {
    $target = $local
}
else {
    $target = Join-Path $env:TEMP 'Invoke-MailboxArchive.ps1'
    $rawUrl = 'https://raw.githubusercontent.com/bluespam-cyber/exchange-auto-archive/main/Invoke-MailboxArchive.ps1'
    Write-Host "[INFO] Invoke-MailboxArchive.ps1 not found next to launcher; downloading from GitHub..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $rawUrl -OutFile $target -UseBasicParsing
}

Write-Host "[INFO] Launcher located script at: $target" -ForegroundColor Cyan

& $target @args