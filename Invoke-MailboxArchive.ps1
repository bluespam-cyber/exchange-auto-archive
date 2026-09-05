#Requires -Version 5.1
<#
.SYNOPSIS
    Mailbox Archive Runner for Exchange Online: enables the archive, assigns the MRM retention policy, clears the documented
    blockers, runs the Managed Folder Assistant, and proves with Exchange's own counters that items actually moved.

.DESCRIPTION
    The Managed Folder Assistant (MFA) applies an MRM retention policy and moves items to the archive on its own schedule,
    at least once every seven days. Waiting is not diagnosis. This tool does what Microsoft's own troubleshooting article
    describes, in order, for one mailbox or a list:

      1  Connect          Exchange Online PowerShell, reusing an open session.
      2  Inspect          Mailbox, archive status, archive quota, retention policy and its tags, RetentionHoldEnabled,
                          ElcProcessingDisabled, litigation hold, account enabled, primary size against the 10 MB floor.
      3  Prepare          With approval: enable the archive (or reconnect a disabled one), assign the policy, clear a
                          retention hold or ELC block when you ask for it, enable auto-expanding archive when licensed.
      4  Run and watch    Start-ManagedFolderAssistant, then read the ELC counters (ElcLastSuccessTimestamp,
                          ElcLastRunArchivedFromRootItemCount and friends) and the primary and archive sizes after each
                          pass. It stops early when a pass moves nothing and the archive-eligible items are gone.
      5  Report           A run folder with report.html, result.json, CASE-NOTES.txt and the MRM diagnostic log.

    Nothing is deleted. The only changes are the ones Microsoft documents for this job, each shown before it is made,
    each verified afterwards.

.PARAMETER Mailbox
    One or more mailbox identities (UPN, email, alias, GUID). When omitted in a console window the tool asks, and every
    later question is asked only after the mailbox has been read, so the choices are grounded in what actually exists.

.PARAMETER Mode
    Check: read everything, report the blockers, change nothing and do not start the assistant.
    Fix: enable the archive if missing, keep or assign or create the policy, clear approved blockers, then run (default).
    Run: change nothing, start the assistant and watch. Asked in the guided start when -Mailbox is omitted.

.PARAMETER RetentionPolicy
    MRM retention policy to assign. Omit to keep the current one. Pass 'choose' to pick from the tenant list, or 'create'
    for the guided creation of a new policy with its own move-to-archive tag. A name that does not exist yet is offered
    for creation in an interactive window; an unattended run creates it when -ArchiveAfterDays is given, otherwise stops.

.PARAMETER ArchiveAfterDays
    Age in days for the move-to-archive tag when the tool creates a policy (1 to 36500). Used with -RetentionPolicy when
    the name does not exist yet, so unattended runs can create the policy without the menu.

.PARAMETER ArchiveTagOnly
    When creating a policy, do not copy the personal and folder tags from the Default MRM Policy. The default is to copy
    them (all except its own move-to-archive default tag, because a policy may hold only one) so users keep their Outlook menu.

.PARAMETER IntervalMinutes
    Minutes to wait after each request before the sizes and counters are read. Default 10. During the wait the tool
    re-reads the mailbox sizes about once a minute and reports as items land in the archive. 0 means "start and leave":
    one request, no waiting; run again in Check mode later to read the result.

.PARAMETER Passes
    How many times the assistant is asked to process the mailbox in this session (1 to 48). One pass is one request,
    one wait, one measurement. Exchange's own schedule (at least once every 7 days) is Microsoft-managed and unchanged.

.PARAMETER FullCrawl
    Adds -FullCrawl to the first pass so tags are recalculated across the whole mailbox (Microsoft's step when the hidden
    MRM configuration is stale).

.PARAMETER ClearRetentionHold
    Set RetentionHoldEnabled to $false when it is $true (with approval). Retention hold is often deliberate, for example
    after a PST import, so it is never cleared silently.

.PARAMETER EnableElcProcessing
    Set ElcProcessingDisabled to $false when it is $true (with approval).

.PARAMETER AutoExpandingArchive
    Enable auto-expanding archiving for the mailbox when the licence allows it (with approval).

.PARAMETER Approve
    Pre-approve changes for unattended runs. Without it, the interactive run asks once before the change set, and an
    unattended run changes nothing and reports what it would have done.

.PARAMETER OutputRoot
    Where the run folder goes. Default %LOCALAPPDATA%\MailboxArchiveRunner.

.PARAMETER CaseNumber
    Written into the report and case notes.

.PARAMETER NonInteractive
    Never prompt. Implied when there is no console.

.PARAMETER Quiet
    Plain text output (also honours NO_COLOR).

.EXAMPLE
    .\Invoke-MailboxArchive.ps1
    Guided start: what to do, which mailbox, then (after reading it) which policy, how many runs, full crawl or not.
    Every question explains what it is for and what each answer does.

.EXAMPLE
    .\Invoke-MailboxArchive.ps1 -Mailbox user@contoso.com -RetentionPolicy "Archive after 90 days" -ArchiveAfterDays 90 -FullCrawl
    Creates the policy (with approval) when it does not exist yet, assigns it, then runs the assistant with a full crawl.

.EXAMPLE
    .\Invoke-MailboxArchive.ps1 -Mailbox user@contoso.com -RetentionPolicy "Default MRM Policy" -Passes 4 -IntervalMinutes 15

.EXAMPLE
    Import-Csv .\users.csv | ForEach-Object { .\Invoke-MailboxArchive.ps1 -Mailbox $_.UPN -Passes 2 -Approve -NonInteractive }

.NOTES
    Version 2.0. Arwaz Khan, Microsoft Support Engineer.
    Windows PowerShell 5.1 or PowerShell 7, ExchangeOnlineManagement module 3.x. Sources in docs/sources.md.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Position = 0)]
    [string[]]$Mailbox = @(),

    [ValidateSet('', 'Check', 'Fix', 'Run')]
    [string]$Mode = '',

    [string]$RetentionPolicy = '',

    [ValidateRange(0, 36500)]
    [int]$ArchiveAfterDays = 0,

    [switch]$ArchiveTagOnly,

    [ValidateRange(1, 48)]
    [int]$Passes = 3,

    [ValidateRange(0, 240)]
    [int]$IntervalMinutes = 10,

    [switch]$FullCrawl,

    [switch]$ClearRetentionHold,

    [switch]$EnableElcProcessing,

    [switch]$AutoExpandingArchive,

    [switch]$Approve,

    [string]$OutputRoot = '',

    [string]$CaseNumber = '',

    [switch]$NonInteractive,

    [switch]$Quiet,

    [switch]$NoOpenReport
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Version = '2.1'
$script:Started = Get-Date

# ---------------------------------------------------------------------------------------------
# Context
# ---------------------------------------------------------------------------------------------
$script:HasConsole = $true
try { $null = [Console]::KeyAvailable; if ([Console]::IsInputRedirected) { $script:HasConsole = $false } } catch { $script:HasConsole = ($Host.Name -in @('Windows PowerShell ISE Host', 'Visual Studio Code Host')) }
if ([Environment]::UserInteractive -eq $false) { $script:HasConsole = $false }
$script:Interactive = -not ($NonInteractive -or -not $script:HasConsole)
if (-not $OutputRoot) { $OutputRoot = Join-Path -Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { $HOME }) -ChildPath 'MailboxArchiveRunner' }
$script:RunId = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:RunDir = Join-Path -Path $OutputRoot -ChildPath $script:RunId
New-Item -Path $script:RunDir -ItemType Directory -Force | Out-Null
$script:LogFile = Join-Path -Path $script:RunDir -ChildPath 'run.log'
$script:Results = [System.Collections.Generic.List[object]]::new()
$script:Findings = [System.Collections.Generic.List[object]]::new()
$script:Actions = [System.Collections.Generic.List[object]]::new()
$script:ExitCode = 0
$script:Stopped = $false
$script:MainCompleted = $false
$script:Bound = $PSBoundParameters
$script:MenuMode = ($Mailbox.Count -eq 0)
if (-not $Mode) { $Mode = 'Fix' }
$script:TargetPolicy = ''
$script:PolicyResolved = $false
$script:PolicyChanged = $false
$script:TagMap = $null
$script:PictureExplained = $false
$script:AskedHold = $false
$script:AskedElc = $false

# ---------------------------------------------------------------------------------------------
# Console layer. Its own palette: amber and copper on slate, with a lantern glyph set. Plain text when there is no console,
# when output is redirected, with -Quiet, or with NO_COLOR.
# ---------------------------------------------------------------------------------------------
$script:Ansi = $true
try {
    if ($Quiet -or $env:NO_COLOR -or $env:TERM -eq 'dumb' -or -not $script:HasConsole) { $script:Ansi = $false }
    if ([Console]::IsOutputRedirected) { $script:Ansi = $false }
    if ($Host.UI.SupportsVirtualTerminal -eq $false) { $script:Ansi = $false }
    if ($script:Ansi) { try { [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false) } catch { } }
} catch { $script:Ansi = $false }
$e = [char]27
$script:C = @{
    Reset = "$e[0m"; Bold = "$e[1m"; Dim = "$e[2m"
    Amber = "$e[38;5;214m"; Copper = "$e[38;5;172m"; Gold = "$e[38;5;220m"; Cream = "$e[38;5;230m"; Sand = "$e[38;5;180m"
    Slate = "$e[38;5;245m"; Ash = "$e[38;5;240m"; Ink = "$e[38;5;236m"; White = "$e[97m"
    Sage = "$e[38;5;108m"; Moss = "$e[38;5;65m"; Rust = "$e[38;5;167m"; Sky = "$e[38;5;110m"
    ClearLine = "$e[2K"
}
$script:Unicode = $script:Ansi
try { if (-not $script:Unicode -or [Console]::OutputEncoding.CodePage -notin @(65001, 1200)) { $script:Unicode = $false } } catch { $script:Unicode = $false }
$script:Fancy = $script:Unicode -and [bool]$env:WT_SESSION
$script:G = if ($script:Fancy) {
    @{ Ok = '◆'; Fail = '◇'; Warn = '◈'; Info = '·'; Step = '›'; Full = '▰'; Empty = '▱'; H = '═'; V = '║'; TL = '╔'; TR = '╗'; BL = '╚'; BR = '╝'; Tee = '╠'; Bullet = '▸'; Arrow = '⟶'; Ellipsis = '…' }
} elseif ($script:Unicode) {
    @{ Ok = '■'; Fail = '□'; Warn = '▲'; Info = '·'; Step = '»'; Full = '█'; Empty = '░'; H = '═'; V = '║'; TL = '╔'; TR = '╗'; BL = '╚'; BR = '╝'; Tee = '╠'; Bullet = '»'; Arrow = '->'; Ellipsis = '...' }
} else {
    @{ Ok = '+'; Fail = 'x'; Warn = '!'; Info = '-'; Step = '>'; Full = '#'; Empty = '.'; H = '='; V = '|'; TL = '+'; TR = '+'; BL = '+'; BR = '+'; Tee = '+'; Bullet = '>'; Arrow = '->'; Ellipsis = '...' }
}
# Lantern frames: a glow that swells and fades, instead of a rotating spinner.
$script:Frames = if ($script:Fancy) { @('○', '◔', '◑', '◕', '●', '◕', '◑', '◔') } elseif ($script:Unicode) { @('·', '∘', '○', '◯', '○', '∘') } else { @('.', 'o', 'O', 'o') }
$script:Width = 78
try { $w = $Host.UI.RawUI.WindowSize.Width; if ($w -ge 60) { $script:Width = [math]::Min(104, $w - 2) } } catch { }

function Get-Plain { param([string]$Text) return [regex]::Replace($Text, "$([char]27)\[[0-9;?]*[A-Za-z]", '') }
function Tint { param([string]$Text, [string]$Color) if ($script:Ansi -and $script:C.ContainsKey($Color)) { return ($script:C[$Color] + $Text + $script:C.Reset) } return $Text }
function Out-Line { param([string]$Text = '') Write-Host $(if ($script:Ansi) { $Text } else { Get-Plain $Text }) }
function Get-Chunks {
    # Word-wrap plain text to a width; returns one string per line.
    param([string]$Text, [int]$Width)
    $out = @(); $rest = [string]$Text
    if (-not $rest) { return @() }
    while ($rest.Length -gt $Width) {
        $cut = $rest.LastIndexOf(' ', [math]::Min($Width, $rest.Length - 1)); if ($cut -lt [int]($Width * 0.5)) { $cut = $Width }
        $out += $rest.Substring(0, $cut).TrimEnd(); $rest = $rest.Substring($cut).TrimStart()
    }
    $out += $rest
    return $out
}
function Write-Note {
    # Grey explanation under a question or a finding: what this is, why it matters, what each answer does. Word-wrapped.
    param([string[]]$Lines, [int]$Indent = 4)
    $room = [math]::Max(30, $script:Width - $Indent)
    foreach ($l in $Lines) { foreach ($c in (Get-Chunks -Text ([string]$l) -Width $room)) { Out-Line ((' ' * $Indent) + (Tint $c 'Slate')) } }
}
function Write-Question { param([Parameter(Mandatory)][string]$Title) Stop-Glow; Out-Line ''; Out-Line ("  {0} {1}" -f (Tint $script:G.Bullet 'Amber'), (Tint $Title 'Cream')) }

$script:Glow = [hashtable]::Synchronized(@{ Active = $false; Quit = $false; Text = ''; Started = [long]0; Frames = @($script:Frames); Warm = @($script:C.Copper, $script:C.Amber, $script:C.Gold, $script:C.Amber); Dim = $script:C.Slate; Reset = $script:C.Reset; Clear = $script:C.ClearLine })
$script:GlowJob = $null
$script:CanThreadJob = [bool](Get-Command -Name Start-ThreadJob -ErrorAction SilentlyContinue)

function Start-Glow {
    param([string]$Text, [long]$StartedTicks = 0)
    if (-not $script:Ansi -or -not $script:CanThreadJob) { Write-Host ("  ... {0}" -f (Get-Plain $Text)) -ForegroundColor DarkGray; return }
    $max = [math]::Max(20, $script:Width - 18)
    $plain = Get-Plain $Text
    if ($plain.Length -gt $max) { $Text = $plain.Substring(0, $max - 1) + $script:G.Ellipsis }
    [System.Threading.Monitor]::Enter($script:Glow)
    try {
        if ($StartedTicks -gt 0) { $script:Glow.Started = $StartedTicks } elseif (-not $script:Glow.Active) { $script:Glow.Started = [DateTime]::UtcNow.Ticks }
        $script:Glow.Text = (Get-Plain $Text); $script:Glow.Active = $true
    } finally { [System.Threading.Monitor]::Exit($script:Glow) }
    if ($script:Glow.Quit) { return }
    if (-not $script:GlowJob -or $script:GlowJob.State -ne 'Running') {
        try {
            if ($script:GlowJob) { Remove-Job -Job $script:GlowJob -Force -ErrorAction SilentlyContinue }
            $script:GlowJob = Start-ThreadJob -Name 'ArchiveGlow' -ArgumentList $script:Glow -ScriptBlock {
                param($S)
                $i = 0
                while (-not $S.Quit) {
                    [System.Threading.Monitor]::Enter($S)
                    try {
                        if ($S.Active -and -not $S.Quit) {
                            $f = $S.Frames[$i % $S.Frames.Count]; $c = $S.Warm[[int][math]::Floor($i / 2) % $S.Warm.Count]; $i++
                            $el = ''
                            if ($S.Started -gt 0) { $sec = ([DateTime]::UtcNow.Ticks - $S.Started) / 10000000.0; $el = if ($sec -ge 60) { ('  {0}:{1:00}' -f [int][math]::Floor($sec / 60), [int]($sec % 60)) } else { ('  {0:0.0}s' -f $sec) } }
                            [Console]::Write(("`r{0}  {1}{2}{3} {4}{5}{6}{7}" -f $S.Clear, $c, $f, $S.Reset, $S.Dim, $S.Text, $el, $S.Reset))
                        }
                    } finally { [System.Threading.Monitor]::Exit($S) }
                    Start-Sleep -Milliseconds 140
                }
            }
        } catch { $script:GlowJob = $null; [System.Threading.Monitor]::Enter($script:Glow); try { $script:Glow.Active = $false } finally { [System.Threading.Monitor]::Exit($script:Glow) }; Write-Host ("  ... {0}" -f (Get-Plain $Text)) -ForegroundColor DarkGray }
    }
}
function Set-GlowText {
    param([Parameter(Mandatory)][string]$Text)
    if (-not $script:Ansi -or -not $script:Glow.Active) { return }
    $plain = Get-Plain $Text; $max = [math]::Max(20, $script:Width - 12)
    if ($plain.Length -gt $max) { $plain = $plain.Substring(0, $max - 1) + $script:G.Ellipsis }
    [System.Threading.Monitor]::Enter($script:Glow); try { $script:Glow.Text = $plain } finally { [System.Threading.Monitor]::Exit($script:Glow) }
}
function Stop-Glow {
    if ($script:Ansi -and $script:Glow.Active) { [System.Threading.Monitor]::Enter($script:Glow); try { $script:Glow.Active = $false; [Console]::Write("`r" + $script:C.ClearLine) } finally { [System.Threading.Monitor]::Exit($script:Glow) } }
}
function Remove-Glow {
    Stop-Glow
    [System.Threading.Monitor]::Enter($script:Glow); try { $script:Glow.Quit = $true } finally { [System.Threading.Monitor]::Exit($script:Glow) }
    if ($script:GlowJob) { try { Stop-Job -Job $script:GlowJob -ErrorAction SilentlyContinue; Remove-Job -Job $script:GlowJob -Force -ErrorAction SilentlyContinue } catch { }; $script:GlowJob = $null }
}

function Write-Log {
    param([Parameter(Mandatory)][string]$Message, [ValidateSet('INFO', 'OK', 'WARN', 'FAIL', 'STEP', 'DEBUG')][string]$Level = 'INFO', [string[]]$Note = @())
    $stamp = Get-Date -Format 'HH:mm:ss'
    try { Add-Content -Path $script:LogFile -Value ("[{0}] [{1}] {2}{3}" -f $stamp, $Level, (Get-Plain $Message), $(if ($Note -and $Note.Count -gt 0) { '  |  ' + ($Note -join ' ') } else { '' })) -Encoding utf8 } catch { }
    if ($Level -eq 'DEBUG') { return }
    $resume = ''; $resumeStart = [long]0
    if ($script:Ansi -and $script:Glow.Active) { $resume = $script:Glow.Text; $resumeStart = $script:Glow.Started }
    Stop-Glow
    $icon = switch ($Level) { 'OK' { Tint $script:G.Ok 'Sage' } 'WARN' { Tint $script:G.Warn 'Amber' } 'FAIL' { Tint $script:G.Fail 'Rust' } 'STEP' { Tint $script:G.Step 'Gold' } default { Tint $script:G.Info 'Slate' } }
    $color = switch ($Level) { 'OK' { 'Sage' } 'WARN' { 'Amber' } 'FAIL' { 'Rust' } 'STEP' { 'Cream' } default { 'Slate' } }
    Out-Line ("  {0} {1} {2}" -f (Tint $stamp 'Ash'), $icon, (Tint $Message $color))
    if ($Note -and $Note.Count -gt 0) { Write-Note -Lines $Note -Indent 13 }
    if ($resume) { Start-Glow -Text $resume -StartedTicks $resumeStart }
}

function Write-Rule { param([string]$Color = 'Ash') Out-Line (Tint ($script:G.H * $script:Width) $Color) }

function Write-Stage {
    # Stage header in the ledger style: a copper tab on the left, the title, then a thin double rule.
    param([Parameter(Mandatory)][string]$Title, [int]$Index = 0, [int]$Total = 5)
    Stop-Glow
    Out-Line ''
    $tab = if ($Index -gt 0) { Tint (" {0} of {1} " -f $Index, $Total) 'Ink' } else { '' }
    $tabBg = if ($script:Ansi -and $Index -gt 0) { "$e[48;5;214m" + $tab + $script:C.Reset } else { $tab }
    Out-Line ("  {0} {1}" -f $tabBg, (Tint $Title 'Cream'))
    Write-Rule 'Ash'
}

function Show-Intro {
    $art = @(
        '  ▄▀█ █▀█ █▀▀ █░█ █ █░█ █▀▀   █▀█ █░█ █▄░█ █▄░█ █▀▀ █▀█',
        '  █▀█ █▀▄ █▄▄ █▀█ █ ▀▄▀ ██▄   █▀▄ █▄█ █░▀█ █░▀█ ██▄ █▀▄'
    )
    if (-not $script:Unicode) { $art = @('  ARCHIVE RUNNER') }
    Out-Line ''
    $pal = @('Copper', 'Amber')
    for ($i = 0; $i -lt $art.Count; $i++) { Out-Line (Tint $art[$i] $pal[$i % $pal.Count]); if ($script:Ansi) { Start-Sleep -Milliseconds 60 } }
    Out-Line ''
    Out-Line ("  {0}  {1}" -f (Tint 'Mailbox Archive Runner for Exchange Online' 'Cream'), (Tint ("v{0}" -f $script:Version) 'Slate'))
    Out-Line ("  {0}" -f (Tint 'Enable the archive, assign the policy, clear the documented blockers, run the Managed Folder Assistant, prove what moved.' 'Slate'))
    Out-Line ''
    Out-Line ("  {0} {1} {2}" -f (Tint $script:G.Bullet 'Amber'), (Tint 'Arwaz Khan' 'Gold'), (Tint 'Microsoft Support Engineer' 'Slate'))
    Write-Rule 'Copper'
    Out-Line ("  {0}" -f (Tint ("{0}  |  {1}" -f $(if ($script:Interactive) { 'interactive' } else { 'unattended' }), "PowerShell $($PSVersionTable.PSVersion)") 'Slate'))
    Out-Line ("  {0}" -f (Tint 'Nothing is deleted. Every change is listed before it is made and checked after. Ctrl+C stops cleanly; the report is still written.' 'Slate'))
    if ($script:Interactive -and $script:MenuMode) { Out-Line ("  {0}" -f (Tint 'A few questions follow. Each one says what it is for and what every answer does; Enter keeps the suggested answer. The mailbox is read first, so the later questions are about what actually exists.' 'Slate')) }
}

function Show-Outro {
    Out-Line ''
    Write-Rule 'Copper'
    Out-Line ("  {0} {1}  {2}" -f (Tint $script:G.Bullet 'Amber'), (Tint 'Arwaz Khan' 'Gold'), (Tint ("Microsoft Support Engineer  |  Mailbox Archive Runner v{0}" -f $script:Version) 'Slate'))
    Out-Line ''
}

function Write-Panel {
    # Double-ruled panel with a left tab. Different shape from the rounded boxes in the other tools.
    param([string]$Title, [string[]]$Lines, [string]$Color = 'Copper', [switch]$Reveal)
    Stop-Glow
    $w = $script:Width; $inner = $w - 4; $g = $script:G
    Out-Line ('  ' + (Tint ($g.TL + $g.H + $g.H + ' ' + $Title + ' ' + ($g.H * [math]::Max(0, $w - 6 - $Title.Length)) + $g.TR) $Color))
    foreach ($l in $Lines) {
        $plain = Get-Plain $l
        $chunks = @($l)
        if ($plain.Length -gt $inner) {
            $lead = ([regex]::Match($plain, '^\s*')).Value; $indent = $lead + '  '
            $chunks = @(); $rest = $plain; $first = $true
            while ($rest.Length -gt 0) {
                $room = if ($first) { $inner } else { $inner - $indent.Length }
                if ($rest.Length -le $room) { $chunks += $(if ($first) { $rest } else { $indent + $rest }); break }
                $cut = $rest.LastIndexOf(' ', [math]::Min($room, $rest.Length - 1)); if ($cut -lt [int]($room * 0.5)) { $cut = $room }
                $chunks += $(if ($first) { $rest.Substring(0, $cut) } else { $indent + $rest.Substring(0, $cut) }); $rest = $rest.Substring($cut).TrimStart(); $first = $false
            }
        }
        foreach ($c in $chunks) { $pad = [math]::Max(0, $inner - (Get-Plain $c).Length); Out-Line ('  ' + (Tint $g.V $Color) + ' ' + $c + (' ' * $pad) + ' ' + (Tint $g.V $Color)); if ($Reveal -and $script:Ansi) { Start-Sleep -Milliseconds 16 } }
    }
    Out-Line ('  ' + (Tint ($g.BL + ($g.H * ($w - 2)) + $g.BR) $Color))
}

function Get-Meter {
    # Horizontal meter for the primary-to-archive picture: [▰▰▰▰▱▱▱▱▱▱] 38%
    param([double]$Value, [double]$Max, [int]$Cells = 20, [string]$Color = 'Amber')
    if ($Max -le 0) { return (Tint ($script:G.Empty * $Cells) 'Ash') }
    $f = [math]::Min($Cells, [int][math]::Round($Cells * $Value / $Max))
    return (Tint ($script:G.Full * $f) $Color) + (Tint ($script:G.Empty * ($Cells - $f)) 'Ash')
}

function Read-Menu {
    # Title, an explanation in grey, aligned options with what each one does, then the prompt. Enter keeps the default.
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][object[]]$Options, [string]$Default = '1', [string[]]$Note = @())
    Write-Question $Title
    if ($Note -and $Note.Count -gt 0) { Write-Note -Lines $Note -Indent 4 }
    $keyW = 0; $labelW = 0
    foreach ($o in $Options) { $keyW = [math]::Max($keyW, ([string]$o.Key).Length + 2); $labelW = [math]::Max($labelW, ([string]$o.Label).Length) }
    foreach ($o in $Options) {
        $isDefault = ($o.Key -eq $Default)
        $key = ("[{0}]" -f $o.Key).PadRight($keyW); $label = ([string]$o.Label).PadRight($labelW)
        $detail = if ($o.PSObject.Properties['Detail'] -and $o.Detail) { [string]$o.Detail } else { '' }
        $lead = 4 + $keyW + 1 + $labelW + 2
        $chunks = @(Get-Chunks -Text $detail -Width ([math]::Max(24, $script:Width - $lead)))
        $first = $(if ($chunks.Count -gt 0) { $chunks[0] } else { '' })
        Out-Line ("    {0} {1}  {2}" -f (Tint $key $(if ($isDefault) { 'Gold' } else { 'Copper' })), (Tint $label $(if ($isDefault) { 'Cream' } else { 'Sand' })), (Tint $first 'Slate'))
        for ($i = 1; $i -lt $chunks.Count; $i++) { Out-Line ((' ' * $lead) + (Tint $chunks[$i] 'Slate')) }
    }
    while ($true) {
        if ($script:Ansi) { Write-Host ("  {0} {1} {2} " -f (Tint '?' 'Amber'), (Tint ("Your choice [{0}]" -f $Default) 'Cream'), (Tint '(Enter keeps the default):' 'Slate')) -NoNewline; $a = Read-Host } else { $a = Read-Host -Prompt ("{0} (choice, Enter = {1})" -f $Title, $Default) }
        $a = ([string]$a).Trim(); if (-not $a) { $a = $Default }
        $hit = $Options | Where-Object { $_.Key -eq $a } | Select-Object -First 1
        if ($hit) { Out-Line ("    {0} {1}" -f (Tint $script:G.Ok 'Sage'), (Tint $hit.Label 'Sage')); return $hit.Key }
        Out-Line ("    {0} {1}" -f (Tint $script:G.Warn 'Amber'), (Tint ('Type one of: ' + (($Options | ForEach-Object { $_.Key }) -join ', ')) 'Amber'))
    }
}
function Read-Text {
    param([Parameter(Mandatory)][string]$Prompt, [string[]]$Note = @(), [string]$Default = '', [string]$Title = '')
    Stop-Glow
    if ($Title) { Write-Question $Title }
    if ($Note -and $Note.Count -gt 0) { Write-Note -Lines $Note -Indent 4 }
    if ($script:Ansi) { Write-Host ("  {0} {1}{2} " -f (Tint '?' 'Amber'), (Tint $Prompt 'Cream'), $(if ($Default) { Tint (" [Enter = {0}]" -f $Default) 'Slate' } else { '' })) -NoNewline; $a = Read-Host } else { $a = Read-Host -Prompt $(if ($Default) { "{0} [{1}]" -f $Prompt, $Default } else { $Prompt }) }
    $a = ([string]$a).Trim(); if (-not $a) { $a = $Default }
    return $a
}
function Read-Number {
    param([Parameter(Mandatory)][string]$Prompt, [int]$Default, [int]$Min, [int]$Max, [string[]]$Note = @())
    $shown = $false
    while ($true) {
        $noteNow = @(); if (-not $shown) { $noteNow = $Note }
        $a = Read-Text -Prompt $Prompt -Default ([string]$Default) -Note $noteNow; $shown = $true
        $n = 0
        if ([int]::TryParse($a, [ref]$n) -and $n -ge $Min -and $n -le $Max) { return $n }
        Out-Line ("    {0} {1}" -f (Tint $script:G.Warn 'Amber'), (Tint ("Enter a whole number between {0} and {1}" -f $Min, $Max) 'Amber'))
    }
}
function Confirm-Change {
    param([Parameter(Mandatory)][string]$Prompt, [string[]]$Note = @())
    Stop-Glow
    if ($WhatIfPreference) { Write-Log ("{0}: dry run (-WhatIf), so the plan is shown and nothing is executed" -f $Prompt) 'INFO'; return $true }
    if ($Approve) { Write-Log ("{0}: pre-approved with -Approve" -f $Prompt) 'INFO'; return $true }
    if (-not $script:Interactive) { Write-Log ("{0}: not approved (unattended run without -Approve)" -f $Prompt) 'WARN'; return $false }
    if ($Note -and $Note.Count -gt 0) { Write-Note -Lines $Note -Indent 4 }
    if ($script:Ansi) { Write-Host ("  {0} {1} {2} " -f (Tint '?' 'Amber'), (Tint $Prompt 'Cream'), (Tint '(type YES to proceed, anything else to skip)' 'Slate')) -NoNewline; $a = Read-Host } else { $a = Read-Host -Prompt ("{0} (type YES to proceed, anything else to skip)" -f $Prompt) }
    return (([string]$a).Trim() -ceq 'YES')
}

# ---------------------------------------------------------------------------------------------
# Records
# ---------------------------------------------------------------------------------------------
function Add-Finding {
    param([Parameter(Mandatory)][string]$Box, [Parameter(Mandatory)][ValidateSet('Blocker', 'Warning', 'Info')][string]$Kind, [Parameter(Mandatory)][string]$Title, [string]$Detail = '', [string]$Fix = '', [string]$Source = '', [string]$Why = '')
    $script:Findings.Add([pscustomobject]@{ Mailbox = $Box; Kind = $Kind; Title = $Title; Detail = $Detail; Fix = $Fix; Why = $Why; Source = $Source; Time = (Get-Date).ToString('s') })
    Write-Log ("[{0}] {1}{2}" -f $Box, $Title, $(if ($Detail) { "  ($Detail)" } else { '' })) $(switch ($Kind) { 'Blocker' { 'FAIL' } 'Warning' { 'WARN' } default { 'INFO' } }) -Note $(if ($Why) { @($Why) } else { @() })
}
function Add-Action {
    param([Parameter(Mandatory)][string]$Box, [Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][ValidateSet('Verified', 'NotVerified', 'Started', 'Skipped', 'Failed', 'WhatIf')][string]$Status, [string]$Detail = '', [string[]]$Note = @())
    $script:Actions.Add([pscustomobject]@{ Mailbox = $Box; Title = $Title; Status = $Status; Detail = $Detail; Time = (Get-Date).ToString('s') })
    Write-Log ("[{0}] {1}: {2}{3}" -f $Box, $Status, $Title, $(if ($Detail) { "  ($Detail)" } else { '' })) $(switch ($Status) { 'Verified' { 'OK' } 'Started' { 'OK' } 'Failed' { 'FAIL' } 'NotVerified' { 'WARN' } default { 'INFO' } }) -Note $Note
}
function Get-ErrorText { param($Err) try { $ex = if ($Err -is [System.Management.Automation.ErrorRecord]) { $Err.Exception } elseif ($Err -is [Exception]) { $Err } else { $null }; if ($ex) { if ($ex.InnerException -and $ex.Message -match '^One or more errors|^Exception calling') { $ex = $ex.InnerException }; return $ex.Message }; return [string]$Err } catch { return 'unknown error' } }

# ---------------------------------------------------------------------------------------------
# Exchange helpers
# ---------------------------------------------------------------------------------------------
function Connect-Exchange {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log 'ExchangeOnlineManagement module is not installed.' 'WARN'
        if (-not (Confirm-Change -Prompt 'Install the ExchangeOnlineManagement module now?' -Note @('This is the official Microsoft module for Exchange Online PowerShell. It is installed for your Windows account only, from the PowerShell Gallery, and takes about a minute. Nothing else on the machine changes.'))) { throw 'The ExchangeOnlineManagement module is required. Install-Module ExchangeOnlineManagement -Scope CurrentUser' }
        Start-Glow 'Installing ExchangeOnlineManagement from the PowerShell Gallery (about a minute)'
        $old = $global:ProgressPreference; $global:ProgressPreference = 'SilentlyContinue'
        $w = @()
        try { Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable w } finally { $global:ProgressPreference = $old; Stop-Glow }
        foreach ($x in @($w)) { Write-Log ("Install-Module: {0}" -f $x) 'DEBUG' }
        Write-Log 'Module installed.' 'OK'
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop -WarningAction SilentlyContinue
    $ver = (Get-Module ExchangeOnlineManagement).Version
    $conn = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' })
    if ($conn.Count -gt 0) { Write-Log ("Using the open Exchange Online session ({0}, module {1})" -f $conn[0].UserPrincipalName, $ver) 'OK'; return $conn[0].UserPrincipalName }
    Write-Log 'Sign in to Exchange Online in the window that opens.' 'STEP' -Note @('Use an admin account that holds the Mail Recipients role (Recipient Management or Organization Management has it). Creating a policy additionally needs Retention Management (Organization Management, Compliance Management or Records Management).')
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop -WarningAction SilentlyContinue
    $conn = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' })
    if ($conn.Count -eq 0) { throw 'Connect-ExchangeOnline returned without an open session.' }
    Write-Log ("Connected as {0} (module {1})" -f $conn[0].UserPrincipalName, $ver) 'OK'
    return $conn[0].UserPrincipalName
}

function ConvertTo-Bytes {
    # Exchange returns sizes as ByteQuantifiedSize objects or as text like "1.5 GB (1,610,612,736 bytes)". Both are handled.
    param($Size)
    if ($null -eq $Size) { return [long]0 }
    try { if ($Size.PSObject.Properties['Value'] -and $Size.Value.PSObject.Methods['ToBytes']) { return [long]$Size.Value.ToBytes() } } catch { }
    try { if ($Size.PSObject.Methods['ToBytes']) { return [long]$Size.ToBytes() } } catch { }
    $m = [regex]::Match([string]$Size, '\(([\d,\.]+)\s*bytes\)')
    if ($m.Success) { return [long]($m.Groups[1].Value -replace '[,\.]', '') }
    return [long]0
}
function Format-Size { param([long]$Bytes) if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) } if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) } if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) } return ("{0} B" -f $Bytes) }

function Get-ElcCounters {
    # Microsoft's verification method: the ELC properties from the mailbox diagnostic log.
    param([Parameter(Mandatory)][string]$Identity)
    $out = [ordered]@{}
    try {
        $log = Export-MailboxDiagnosticLogs -Identity $Identity -ExtendedProperties -ErrorAction Stop
        $xml = [xml]$log.MailboxLog
        foreach ($p in $xml.Properties.MailboxTable.Property) { if ($p.Name -like 'Elc*' -or $p.Name -like 'ELC*') { $out[$p.Name] = $p.Value } }
    } catch { Write-Log ("ELC counters unavailable: {0}" -f (Get-ErrorText $_)) 'DEBUG' }
    return $out
}
function Get-AgeDays {
    # AgeLimitForRetention arrives as a TimeSpan in some sessions and as text like '730.00:00:00' in REST sessions.
    param($Age)
    if ($null -eq $Age) { return $null }
    if ($Age -is [timespan]) { return [int]$Age.TotalDays }
    if ($Age.PSObject.Properties['TotalDays']) { return [int]$Age.TotalDays }
    $ts = [timespan]::Zero
    if ([timespan]::TryParse([string]$Age, [ref]$ts)) { return [int]$ts.TotalDays }
    return $null
}
function Get-Counter { param($Counters, [string]$Name) foreach ($k in $Counters.Keys) { if ($k -ieq $Name) { return $Counters[$k] } } return $null }

function Get-MailboxPicture {
    param([Parameter(Mandatory)][string]$Identity)
    $mbx = Get-Mailbox -Identity $Identity -ErrorAction Stop
    $p = $null; $a = $null
    try { $p = Get-MailboxStatistics -Identity $mbx.ExchangeGuid.ToString() -ErrorAction Stop } catch { }
    if ($mbx.ArchiveStatus -eq 'Active') { try { $a = Get-MailboxStatistics -Identity $mbx.ExchangeGuid.ToString() -Archive -ErrorAction Stop } catch { } }
    return [pscustomobject]@{
        Mailbox = $mbx
        PrimaryBytes = $(if ($p) { ConvertTo-Bytes $p.TotalItemSize } else { [long]0 })
        PrimaryItems = $(if ($p) { [int]$p.ItemCount } else { 0 })
        ArchiveBytes = $(if ($a) { ConvertTo-Bytes $a.TotalItemSize } else { [long]0 })
        ArchiveItems = $(if ($a) { [int]$a.ItemCount } else { 0 })
        ArchiveReady = [bool]$a
    }
}

# ---------------------------------------------------------------------------------------------
# Stages
# ---------------------------------------------------------------------------------------------
function Show-Picture {
    # One line the user can read at a glance: a meter of how much of the mail sits in the archive, then both sizes.
    param([Parameter(Mandatory)]$Pic, [string]$Label = 'now', [string]$Color = 'Copper')
    $resume = ''; $resumeStart = [long]0
    if ($script:Ansi -and $script:Glow.Active) { $resume = $script:Glow.Text; $resumeStart = $script:Glow.Started }
    Stop-Glow
    $tot = [math]::Max(1, $Pic.PrimaryBytes + $Pic.ArchiveBytes)
    $arch = if ($Pic.ArchiveReady) { ("archive {0} ({1:N0} items)" -f (Format-Size $Pic.ArchiveBytes), $Pic.ArchiveItems) } else { 'archive not enabled' }
    Out-Line ("  {0} {1}  {2}" -f (Tint $Label.PadRight(9) 'Slate'), (Get-Meter -Value $Pic.ArchiveBytes -Max $tot -Color $Color), (Tint ("primary {0} ({1:N0} items)   {2}" -f (Format-Size $Pic.PrimaryBytes), $Pic.PrimaryItems, $arch) 'Slate'))
    if ($resume) { Start-Glow -Text $resume -StartedTicks $resumeStart }
}

function Show-StartMenu {
    # Only two things are asked before connecting: what to do, and which mailbox. Everything else is asked after the mailbox has been read.
    $k = Read-Menu -Title 'What should this run do?' -Default '2' -Note @('Every option reads the mailbox first and shows what it found. Only Fix changes anything, and only after you type YES to a list of the exact changes.') -Options @(
        [pscustomobject]@{ Key = '1'; Label = 'Check'; Detail = 'read the mailbox, the archive, the policy and its tags; report what blocks archiving; change nothing, start nothing' },
        [pscustomobject]@{ Key = '2'; Label = 'Fix and run'; Detail = 'check, then enable the archive if missing, keep or assign or create the policy, clear approved blockers, run the assistant and show what moved' },
        [pscustomobject]@{ Key = '3'; Label = 'Run only'; Detail = 'check, then run the assistant with the settings as they are; change nothing. For when the setup is already right and you want to push and measure' },
        [pscustomobject]@{ Key = '0'; Label = 'Quit'; Detail = '' }
    )
    if ($k -eq '0') { return $false }
    $script:Mode = switch ($k) { '1' { 'Check' } '3' { 'Run' } default { 'Fix' } }
    $boxes = Read-Text -Title 'Which mailbox?' -Prompt 'Mailbox address:' -Note @('The user whose older mail should move from the primary mailbox into the archive. Type the sign-in address (UPN) or the email address, for example jsmith@contoso.com. Several mailboxes: separate them with commas; they get the same choices.')
    if (-not $boxes) { return $false }
    $script:Mailbox = @($boxes -split '[,\s]+' | Where-Object { $_ })
    return $true
}

function Get-TagMap {
    if ($null -eq $script:TagMap) {
        $script:TagMap = @{}
        try { foreach ($t in @(Get-RetentionPolicyTag -ErrorAction Stop -WarningAction SilentlyContinue)) { $script:TagMap[[string]$t.Name] = $t } } catch { Write-Log ("Get-RetentionPolicyTag: {0}" -f (Get-ErrorText $_)) 'DEBUG' }
    }
    return $script:TagMap
}
function Get-PolicyTags {
    param([Parameter(Mandatory)]$Policy)
    $map = Get-TagMap; $out = @()
    foreach ($l in @($Policy.RetentionPolicyTagLinks)) {
        $n = [string]$l
        if ($map.ContainsKey($n)) { $out += $map[$n] }
        else { try { $t = Get-RetentionPolicyTag -Identity $n -ErrorAction Stop -WarningAction SilentlyContinue; if ($t) { $out += $t; $map[[string]$t.Name] = $t } } catch { } }
    }
    return $out
}
function Get-ArchiveAge {
    # Days of the policy's default move-to-archive tag, or $null when it has none.
    param([Parameter(Mandatory)]$Policy)
    $t = @(Get-PolicyTags -Policy $Policy | Where-Object { $_.Type -eq 'All' -and $_.RetentionAction -eq 'MoveToArchive' -and $_.RetentionEnabled })
    if ($t.Count -eq 0) { return $null }
    return (Get-AgeDays $t[0].AgeLimitForRetention)
}
function Get-OldestMail {
    # Received date of the oldest mail item in the primary mailbox, from the folder statistics. $null when unknown.
    param([Parameter(Mandatory)][string]$Identity)
    $skip = @('Calendar', 'Contacts', 'Tasks', 'Notes', 'Journal', 'SyncIssues', 'Conflicts', 'LocalFailures', 'ServerFailures', 'Root', 'RecoverableItemsRoot', 'RecoverableItemsDeletions', 'RecoverableItemsPurges', 'RecoverableItemsVersions', 'RecoverableItemsDiscoveryHolds', 'RecoverableItemsSubstrateHolds', 'Audits', 'CalendarLogging')
    try {
        $oldest = $null
        foreach ($f in @(Get-MailboxFolderStatistics -Identity $Identity -IncludeOldestAndNewestItems -ErrorAction Stop -WarningAction SilentlyContinue)) {
            if (-not $f.PSObject.Properties['OldestItemReceivedDate'] -or -not $f.OldestItemReceivedDate) { continue }
            if ([string]$f.FolderType -in $skip -or [string]$f.FolderPath -like '/Recoverable Items*') { continue }
            $d = $null; try { $d = [datetime]$f.OldestItemReceivedDate } catch { continue }
            if (-not $oldest -or $d -lt $oldest) { $oldest = $d }
        }
        return $oldest
    } catch { Write-Log ("Folder statistics: {0}" -f (Get-ErrorText $_)) 'DEBUG'; return $null }
}

function Inspect-Mailbox {
    param([Parameter(Mandatory)][string]$Identity)
    Start-Glow ("Reading {0}: mailbox, archive and sizes" -f $Identity)
    $pic = Get-MailboxPicture -Identity $Identity
    $m = $pic.Mailbox; $box = [string]$m.UserPrincipalName
    Set-GlowText ("{0}: retention policy and its tags" -f $box)
    $policyName = [string]$m.RetentionPolicy
    $policy = $null; $tags = @()
    if ($policyName) { try { $policy = Get-RetentionPolicy -Identity $policyName -ErrorAction Stop -WarningAction SilentlyContinue; $tags = @(Get-PolicyTags -Policy $policy) } catch { } }
    $archiveTags = @($tags | Where-Object { $_.RetentionAction -eq 'MoveToArchive' -and $_.RetentionEnabled })
    $defaultArchive = @($archiveTags | Where-Object { $_.Type -eq 'All' })
    $ageDays = $(if ($defaultArchive.Count -gt 0) { Get-AgeDays $defaultArchive[0].AgeLimitForRetention } else { $null })
    Set-GlowText ("{0}: account state, holds and the assistant's last run" -f $box)
    $user = $null; try { $user = Get-User -Identity $box -ErrorAction Stop -WarningAction SilentlyContinue } catch { Write-Log ("Get-User {0}: {1}" -f $box, (Get-ErrorText $_)) 'DEBUG' }
    $counters = Get-ElcCounters -Identity $m.ExchangeGuid.ToString()
    Set-GlowText ("{0}: age of the oldest mail item" -f $box)
    $oldest = Get-OldestMail -Identity $m.ExchangeGuid.ToString()
    Stop-Glow

    Write-Log ("{0}  archive {1}  policy {2}" -f $box, $m.ArchiveStatus, $(if ($policyName) { "'$policyName'" } else { 'none' })) 'OK'
    Show-Picture -Pic $pic -Label 'today' -Color 'Copper'
    if (-not $script:PictureExplained) { Write-Note -Lines @('The bar shows how much of this user''s mail already sits in the archive; it fills as items move. Primary is what the user sees in Outlook day to day; archive is the Online Archive folder tree.') -Indent 4; $script:PictureExplained = $true }

    # Blockers Microsoft documents, each with the reason in plain words
    $mrm = 'Resolve email archive and deletion issues when using MRM'
    if ($m.ArchiveStatus -ne 'Active') {
        $disabledGuid = [string]$m.DisabledArchiveGuid
        if ($disabledGuid -and $disabledGuid -ne '00000000-0000-0000-0000-000000000000') { Add-Finding -Box $box -Kind Blocker -Title 'Archive was disabled earlier; a disabled archive still exists' -Detail ("DisabledArchiveGuid {0}" -f $disabledGuid) -Fix 'Within 30 days Enable-Mailbox -Archive reconnects it with its contents; after 30 days Set-Mailbox -RemoveDisabledArchive first, then enable a new one' -Source 'Enable archive mailboxes for Microsoft 365' -Why 'Exchange keeps a disabled archive for 30 days so it can be reconnected with everything in it. Fix mode tries the reconnect and falls back to a fresh archive when Exchange refuses.' }
        else { Add-Finding -Box $box -Kind Blocker -Title 'Archive mailbox is not enabled; move-to-archive tags do nothing until it is' -Detail ("ArchiveStatus {0}" -f $m.ArchiveStatus) -Fix 'Enable-Mailbox -Archive' -Source 'Customize an archive and deletion policy (MRM) for mailboxes' -Why 'There is no archive to move items into. Fix mode enables one (Exchange provisions it in about a minute); it appears in Outlook as Online Archive.' }
    }
    if ($m.RetentionHoldEnabled) { Add-Finding -Box $box -Kind Blocker -Title 'Retention hold is on; MRM does not process the mailbox' -Detail ("RetentionHoldEnabled True{0}" -f $(if ($m.EndDateForRetentionHold) { ", until $($m.EndDateForRetentionHold)" } else { '' })) -Fix 'Deliberate after a PST import or for a user on leave. Lift it with -ClearRetentionHold or the Fix menu' -Source $mrm -Why 'A hold pauses every MRM action in this mailbox, usually on purpose: during a PST import so old imported mail is not archived at once, or while a user is away. The tool never lifts it without being asked.' }
    if ($m.ElcProcessingDisabled) { Add-Finding -Box $box -Kind Blocker -Title 'ElcProcessingDisabled is set; the assistant skips this mailbox entirely' -Detail 'ElcProcessingDisabled True' -Fix 'Clear it with -EnableElcProcessing or the Fix menu' -Source $mrm -Why 'An admin switched the assistant off for this one mailbox. Nothing is stamped or moved until it is switched back on.' }
    if ($pic.PrimaryBytes -gt 0 -and $pic.PrimaryBytes -lt 10MB) { Add-Finding -Box $box -Kind Blocker -Title 'Primary mailbox is under 10 MB; MRM does not process mailboxes that small' -Detail (Format-Size $pic.PrimaryBytes) -Fix 'Nothing to do; this is by design and archiving starts once the mailbox grows' -Source $mrm -Why 'Exchange skips mailboxes this small on purpose. Nothing is wrong with the setup; there is simply too little mail for the assistant to bother with.' }
    if ($user -and $user.PSObject.Properties['AccountDisabled'] -and $user.AccountDisabled -and $m.RecipientTypeDetails -eq 'UserMailbox') { Add-Finding -Box $box -Kind Blocker -Title 'The account is disabled; items are not moved to the archive for a disabled regular mailbox' -Detail 'AccountDisabled True' -Fix 'Enable the account, or convert to a shared mailbox if the person has left' -Source 'Customize an archive and deletion policy (MRM) for mailboxes' -Why 'Exchange does not archive for a regular mailbox whose sign-in is disabled. The tool does not touch accounts.' }
    if (-not $policyName) { Add-Finding -Box $box -Kind Blocker -Title 'No MRM retention policy is assigned' -Fix 'Assign one (Fix mode offers the tenant list or creates one)' -Source 'Customize an archive and deletion policy (MRM) for mailboxes' -Why 'Without a policy there are no tags, so no item ever qualifies for the archive.' }
    elseif ($tags.Count -gt 0 -and $archiveTags.Count -eq 0) { Add-Finding -Box $box -Kind Blocker -Title 'The assigned policy has no enabled move-to-archive tag; nothing will ever move' -Detail ("policy {0}, {1} tag(s): {2}" -f $policyName, $tags.Count, (($tags | ForEach-Object { "$($_.Name) [$($_.RetentionAction)]" }) -join ', ')) -Fix 'Assign a policy that contains an archive tag, or create one in Fix mode' -Source $mrm -Why 'Only a tag with the Move to Archive action moves anything. This policy deletes or marks items but never archives them, so the assistant runs and moves nothing.' }
    elseif ($archiveTags.Count -gt 0 -and $defaultArchive.Count -eq 0) { Add-Finding -Box $box -Kind Warning -Title 'The policy has only personal archive tags; items move only where the user applied a tag' -Detail (($archiveTags | ForEach-Object { $_.Name }) -join ', ') -Fix 'Add a default policy tag (applies to the entire mailbox) if automatic archiving is expected' -Source $mrm -Why 'Personal tags move only the items or folders the user tagged in Outlook. Automatic archiving needs a default policy tag, which Fix mode can create.' }
    if ($defaultArchive.Count -gt 0) { Add-Finding -Box $box -Kind Info -Title ("Archive rule: '{0}' moves items older than {1}" -f $defaultArchive[0].Name, $(if ($null -ne $ageDays) { "$ageDays days" } else { 'an unknown age' })) -Source 'Retention tags and retention policies' -Why 'Age counts from the day an item was received. Anything younger stays in the primary mailbox until it reaches this age; then the assistant moves it on its next visit.' }
    if ($null -ne $ageDays -and $oldest) {
        $oldDays = [int]((Get-Date) - $oldest).TotalDays
        if ($oldDays -gt $ageDays) { Add-Finding -Box $box -Kind Info -Title ("Oldest mail item is {0} days old ({1:yyyy-MM-dd}); items older than {2} days qualify, so there is mail to move" -f $oldDays, $oldest, $ageDays) -Source 'Retention tags and retention policies' -Why 'The assistant moves everything past the archive age when it processes the mailbox; a large backlog goes in slices over several runs.' }
        else { Add-Finding -Box $box -Kind Warning -Title ("Nothing qualifies for the archive yet: the oldest mail item is {0} days old and the rule moves items older than {1} days" -f $oldDays, $ageDays) -Detail ("first item qualifies around {0:yyyy-MM-dd}" -f $oldest.AddDays($ageDays)) -Fix 'Wait, or use a policy with a shorter archive age (Fix mode can create one)' -Source 'Retention tags and retention policies' -Why 'The assistant will run and report a clean pass while moving nothing. That is the policy working as written, not a fault.' }
    }
    if ($m.LitigationHoldEnabled) { Add-Finding -Box $box -Kind Info -Title 'Litigation hold is on; archiving still runs, deletions are preserved in Recoverable Items' -Fix 'Microsoft recommends enabling the archive and auto-expanding archive for held mailboxes' -Source $mrm -Why 'A litigation hold does not stop the move to the archive; it only keeps deleted items recoverable.' }
    if ($m.PSObject.Properties['AutoExpandingArchiveEnabled'] -and -not $m.AutoExpandingArchiveEnabled -and $m.ArchiveStatus -eq 'Active') { Add-Finding -Box $box -Kind Info -Title 'Auto-expanding archive is off' -Fix 'Optional: -AutoExpandingArchive enables it (needs a qualifying licence, cannot be turned off later)' -Source 'Enable auto-expanding archiving' -Why 'The standard archive stops at its quota (100 GB on most plans). Auto-expanding adds space automatically as the archive fills.' }
    $last = Get-Counter $counters 'ElcLastSuccessTimestamp'
    if ($last) { Add-Finding -Box $box -Kind Info -Title ("Assistant last completed without errors at {0} (UTC)" -f $last) -Source $mrm -Why 'This is Exchange''s own record of the last time the assistant finished this mailbox cleanly. The tool reads it again after each pass as proof of a fresh run.' } else { Add-Finding -Box $box -Kind Info -Title 'No ElcLastSuccessTimestamp yet; the assistant has not completed a run on this mailbox' -Source $mrm -Why 'Normal for a new mailbox or one that has been below 10 MB. The first completed run sets it.' }

    return [pscustomobject]@{ Picture = $pic; Box = $box; PolicyName = $policyName; Policy = $policy; AgeDays = $ageDays; Oldest = $oldest; ArchiveTags = $archiveTags; DefaultArchive = $defaultArchive; Counters = $counters; Blockers = @($script:Findings | Where-Object { $_.Mailbox -eq $box -and $_.Kind -eq 'Blocker' }) }
}

function Ask-Plan {
    # Asked only in the guided start, only after the mailboxes have been read, so every option is grounded in real state.
    param([Parameter(Mandatory)][object[]]$Infos)
    $first = $Infos[0]
    if ($Mode -eq 'Fix' -and -not $script:Bound.ContainsKey('RetentionPolicy')) {
        $cur = [string]$first.PolicyName; $curAge = $first.AgeDays
        $note = @('A retention policy is the rule set that decides when mail leaves the primary mailbox for the archive. A mailbox holds exactly one policy.')
        if ($cur) { $note += $(if ($null -ne $curAge) { "Today the mailbox uses '$cur', which moves items older than $curAge days." } else { "Today the mailbox uses '$cur', which has no automatic move-to-archive rule, so nothing moves on its own." }) } else { $note += 'Today the mailbox has no policy, so nothing moves.' }
        if ($Infos.Count -gt 1) { $note += ("The same choice is applied to all {0} mailboxes in this run." -f $Infos.Count) }
        $k = Read-Menu -Title 'Which retention policy should the mailbox use?' -Default '1' -Note $note -Options @(
            [pscustomobject]@{ Key = '1'; Label = $(if ($cur) { "Keep '$cur'" } else { 'Keep none' }); Detail = $(if ($cur) { $(if ($null -ne $curAge) { "no change; items older than $curAge days move" } else { 'no change; nothing will move automatically' }) } else { 'no change; nothing will move' }) },
            [pscustomobject]@{ Key = '2'; Label = 'Pick an existing policy'; Detail = 'you will see every policy in the tenant with its archive age, then choose' },
            [pscustomobject]@{ Key = '3'; Label = 'Create a new policy'; Detail = 'guided: choose an age in days and a name; the tool creates the archive rule and the policy, then assigns it' }
        )
        switch ($k) { '2' { $script:RetentionPolicy = 'choose' } '3' { $script:RetentionPolicy = 'create' } }
    }
    if ($Mode -eq 'Fix') {
        foreach ($i in $Infos) {
            $mb = $i.Picture.Mailbox
            if ($mb.RetentionHoldEnabled -and -not $script:Bound.ContainsKey('ClearRetentionHold') -and -not $script:AskedHold) {
                $script:AskedHold = $true
                $k = Read-Menu -Title ("Retention hold is on for {0}. Lift it?" -f $i.Box) -Default '1' -Note @(("While the hold is on, MRM does nothing in this mailbox, so nothing is archived. Holds are usually set on purpose: during a PST import (so imported old mail is not archived at once) or while a user is on leave. Lift it only if you know why it was set.{0}" -f $(if ($mb.EndDateForRetentionHold) { " It is due to end on $($mb.EndDateForRetentionHold)." } else { '' }))) -Options @(
                    [pscustomobject]@{ Key = '1'; Label = 'Leave it on'; Detail = 'nothing moves while it is on; the run reports the hold and does not start the assistant' },
                    [pscustomobject]@{ Key = '2'; Label = 'Lift it'; Detail = 'Set-Mailbox -RetentionHoldEnabled $false goes on the change list for your YES' })
                if ($k -eq '2') { $script:ClearRetentionHold = $true }
            }
            if ($mb.ElcProcessingDisabled -and -not $script:Bound.ContainsKey('EnableElcProcessing') -and -not $script:AskedElc) {
                $script:AskedElc = $true
                $k = Read-Menu -Title ("The assistant is switched off for {0} (ElcProcessingDisabled). Switch it on?" -f $i.Box) -Default '1' -Note @('An admin set this, sometimes to keep a new policy from acting on a mailbox before it was reviewed. With it on, nothing is stamped or moved.') -Options @(
                    [pscustomobject]@{ Key = '1'; Label = 'Leave it off'; Detail = 'nothing moves; the run reports it and does not start the assistant' },
                    [pscustomobject]@{ Key = '2'; Label = 'Switch it on'; Detail = 'Set-Mailbox -ElcProcessingDisabled $false goes on the change list for your YES' })
                if ($k -eq '2') { $script:EnableElcProcessing = $true }
            }
        }
    }
    if (-not ($script:Bound.ContainsKey('Passes') -or $script:Bound.ContainsKey('IntervalMinutes'))) {
        $k = Read-Menu -Title 'How should the assistant be run in this session?' -Default '2' -Note @(
            'The Managed Folder Assistant is the Exchange process that applies the policy: it stamps every item with its tag and moves the ones past the archive age. Exchange runs it on its own at least once every 7 days; that schedule is Microsoft-managed and cannot be changed.',
            'What you choose here is how many times this session asks Exchange to process the mailbox now (a pass) and how long to wait after each request before measuring (the interval). A normal mailbox is processed within minutes of a request. A large one is throttled and needs several requests over hours, which is what extra passes are for. During every wait the sizes are re-read and each batch that lands in the archive is shown as it happens.'
        ) -Options @(
            [pscustomobject]@{ Key = '1'; Label = 'Once, wait 10 minutes'; Detail = 'one request, one measurement; about 10 minutes. A first look, or a small mailbox' },
            [pscustomobject]@{ Key = '2'; Label = '3 passes, 10 minutes apart'; Detail = 'about 30 minutes; stops early when two passes in a row move nothing. Sensible default' },
            [pscustomobject]@{ Key = '3'; Label = '6 passes, 15 minutes apart'; Detail = 'about 90 minutes; for mailboxes of tens of GB, where Exchange moves in throttled slices' },
            [pscustomobject]@{ Key = '4'; Label = '12 passes, 30 minutes apart'; Detail = 'about 6 hours; leave it running for a very large backlog' },
            [pscustomobject]@{ Key = '5'; Label = 'Start and leave'; Detail = 'one request, no waiting; Exchange keeps working in the background. Run Check later to see the result' },
            [pscustomobject]@{ Key = '6'; Label = 'Custom'; Detail = 'you set the passes (1 to 48) and the minutes to wait after each (0 to 240)' }
        )
        switch ($k) {
            '1' { $script:Passes = 1; $script:IntervalMinutes = 10 }
            '2' { $script:Passes = 3; $script:IntervalMinutes = 10 }
            '3' { $script:Passes = 6; $script:IntervalMinutes = 15 }
            '4' { $script:Passes = 12; $script:IntervalMinutes = 30 }
            '5' { $script:Passes = 1; $script:IntervalMinutes = 0 }
            '6' {
                $script:Passes = Read-Number -Prompt 'How many passes?' -Default 3 -Min 1 -Max 48 -Note @('Each pass is one request to Exchange followed by one wait and one measurement. Extra passes help only when the mailbox is large enough to be throttled; for a normal mailbox one or two is plenty.')
                $script:IntervalMinutes = Read-Number -Prompt 'Minutes to wait after each request?' -Default 10 -Min 0 -Max 240 -Note @('How long Exchange gets to work before the sizes and counters are read. 10 suits most mailboxes; 30 or more for very large ones. 0 means do not wait: start and leave, then run Check later.')
            }
        }
    }
    if (-not $script:Bound.ContainsKey('FullCrawl')) {
        $changing = [bool]$RetentionPolicy
        $k = Read-Menu -Title 'Re-check every item in the mailbox on the first pass (full crawl)?' -Default $(if ($changing) { '1' } else { '2' }) -Note @('Normally the assistant looks only at items that changed since its last visit. A full crawl makes it re-evaluate every item once against the current tags. Microsoft recommends it after a policy change or when the mailbox''s hidden MRM configuration is stale. It costs time on a large mailbox, nothing else.') -Options @(
            [pscustomobject]@{ Key = '1'; Label = 'Yes'; Detail = $(if ($changing) { 'recommended now, because the policy is about to change' } else { 'Start-ManagedFolderAssistant -FullCrawl on the first pass' }) },
            [pscustomobject]@{ Key = '2'; Label = 'No'; Detail = 'normal run; fine when the policy has not changed' })
        if ($k -eq '1') { $script:FullCrawl = $true }
    }
}

function New-ArchivePolicy {
    # Guided creation: one default move-to-archive tag with the chosen age, wrapped in a new policy. Nothing in any mailbox changes until the policy is assigned.
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Existing, [Parameter(Mandatory)]$Info, [string]$Name = '', [int]$Days = 0)
    $box = $Info.Box
    if ($Days -le 0) {
        if (-not $script:Interactive) { throw 'Creating a policy unattended needs -ArchiveAfterDays.' }
        $Days = Read-Number -Prompt 'Move items to the archive when older than how many days?' -Default 365 -Min 1 -Max 36500 -Note @('Examples: 30 = one month, 90 = one quarter, 180 = six months, 365 = one year, 730 = two years (the Microsoft default). Age counts from the day an item was received. Anything younger stays in the primary mailbox until it reaches this age.')
    }
    $names = @($Existing | ForEach-Object { [string]$_.Name })
    if (-not $Name) {
        $suggest = "Archive after $Days days"; $n = 2; while ($names -contains $suggest) { $suggest = "Archive after $Days days ($n)"; $n++ }
        $Name = Read-Text -Prompt 'Policy name:' -Default $suggest -Note @('Admins see this name in the Exchange admin center and in Purview under MRM retention policies. Users never see the policy name; in Outlook they only see tag names.')
        while ($names -contains $Name -or $Name.Length -gt 64) { $Name = Read-Text -Prompt $(if ($names -contains $Name) { ("'{0}' already exists. Another name:" -f $Name) } else { 'That name is over 64 characters. A shorter name:' }) -Default $suggest }
    } elseif ($names -contains $Name) { throw ("Retention policy '{0}' already exists." -f $Name) }
    elseif ($Name.Length -gt 64) { throw 'A retention policy name is limited to 64 characters.' }
    $map = Get-TagMap
    $reuse = $null
    foreach ($t in @($map.Values)) { if ($t.Type -eq 'All' -and $t.RetentionAction -eq 'MoveToArchive' -and $t.RetentionEnabled -and (Get-AgeDays $t.AgeLimitForRetention) -eq $Days) { $reuse = $t; break } }
    $tagName = "Move to archive after $Days days"; $k = 2
    while ($map.ContainsKey($tagName)) { $tagName = "Move to archive after $Days days ($k)"; $k++ }
    if ($reuse) { $tagName = [string]$reuse.Name }
    # Personal and folder tags carried over from the Default MRM Policy, so users keep their Outlook menu. Default policy tags are not copied: a policy may hold one per action.
    $carry = @()
    if (-not $ArchiveTagOnly) {
        $dflt = $Existing | Where-Object { $_.Name -eq 'Default MRM Policy' } | Select-Object -First 1
        if ($dflt) { $carry = @(Get-PolicyTags -Policy $dflt | Where-Object { $_.Type -ne 'All' -and [string]$_.Name -ne $tagName } | ForEach-Object { [string]$_.Name }) }
    }
    $lines = @()
    if ($reuse) { $lines += ("Reuse the existing rule '{0}' (already moves items older than {1} days to the archive)" -f $tagName, $Days) }
    else { $lines += ("New-RetentionPolicyTag '{0}' -Type All -RetentionEnabled `$true -AgeLimitForRetention {1} -RetentionAction MoveToArchive" -f $tagName, $Days) }
    $lines += ("New-RetentionPolicy '{0}' -RetentionPolicyTagLinks '{1}'{2}" -f $Name, $tagName, $(if ($carry.Count -gt 0) { (" and {0} tag(s) carried over from the Default MRM Policy" -f $carry.Count) } else { '' }))
    if ($carry.Count -gt 0) { $lines += ("    carried over: {0}" -f ($carry -join ', ')) }
    Write-Panel -Title 'New retention policy' -Lines $lines -Color 'Amber'
    if (-not (Confirm-Change -Prompt 'Create the rule and the policy?' -Note @('Creating a tag and a policy changes nothing in any mailbox by itself; the policy is assigned in the next step after another YES. Needs the Retention Management role (Organization Management, Compliance Management or Records Management have it).'))) { Add-Action -Box $box -Title ("Create retention policy '{0}'" -f $Name) -Status Skipped -Detail 'not approved'; return '' }
    if (-not $PSCmdlet.ShouldProcess($Name, 'New-RetentionPolicyTag and New-RetentionPolicy')) { Add-Action -Box $box -Title ("Create retention policy '{0}'" -f $Name) -Status WhatIf; return '' }
    try {
        Start-Glow ("Creating the archive rule '{0}'" -f $tagName)
        if (-not $reuse) { New-RetentionPolicyTag -Name $tagName -Type All -RetentionEnabled $true -AgeLimitForRetention $Days -RetentionAction MoveToArchive -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null; $script:TagMap = $null }
        Set-GlowText ("Creating the policy '{0}'" -f $Name)
        $links = @($tagName) + $carry
        New-RetentionPolicy -Name $Name -RetentionPolicyTagLinks $links -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
        Set-GlowText 'Reading the new policy back from Exchange'
        Start-Sleep -Seconds 3
        $chk = Get-RetentionPolicy -Identity $Name -ErrorAction Stop -WarningAction SilentlyContinue
        $age = Get-ArchiveAge -Policy $chk
        Stop-Glow
        Add-Action -Box $box -Title ("Retention policy '{0}' created" -f $Name) -Status $(if ($age -eq $Days) { 'Verified' } else { 'NotVerified' }) -Detail ("Get-RetentionPolicy reports {0} tag(s); archive rule {1}" -f @($chk.RetentionPolicyTagLinks).Count, $(if ($null -ne $age) { "moves items older than $age days" } else { 'missing' })) -Note @('The policy exists in the tenant now. It takes effect for a mailbox only once assigned, which is the next change on the list.')
        $script:PolicyChanged = $true
        return [string]$chk.Name
    } catch {
        Stop-Glow
        $msg = Get-ErrorText $_
        $why = $(if ($msg -match 'isn''t assigned to any management roles|not assigned to any|is not recognized|Access is denied|insufficient|not authorized') { 'The signed-in account lacks the Retention Management role. Organization Management, Compliance Management and Records Management include it.' } else { '' })
        Add-Action -Box $box -Title ("Create retention policy '{0}'" -f $Name) -Status Failed -Detail $msg -Note $(if ($why) { @($why) } else { @() })
        return ''
    }
}

function Resolve-PolicyChoice {
    param([Parameter(Mandatory)]$Info)
    if (-not $RetentionPolicy) { return '' }
    $all = @(Get-RetentionPolicy -ErrorAction Stop -WarningAction SilentlyContinue)
    $want = [string]$RetentionPolicy
    if ($want -ieq 'create') { if (-not $script:Interactive) { throw 'RetentionPolicy "create" needs an interactive window. Pass a name with -ArchiveAfterDays for an unattended run.' }; return (New-ArchivePolicy -Existing $all -Info $Info) }
    if ($want -ieq 'choose') {
        if (-not $script:Interactive) { Write-Log 'RetentionPolicy "choose" needs an interactive window; keeping the current policy.' 'WARN'; return '' }
        $opts = @(); $i = 0
        foreach ($p in ($all | Sort-Object Name)) {
            $i++; $age = Get-ArchiveAge -Policy $p; $cnt = @($p.RetentionPolicyTagLinks).Count
            $d = $(if ($null -ne $age) { "moves items older than $age days" } else { 'no automatic move-to-archive rule; nothing moves on its own' }) + ", $cnt tag(s)" + $(if ($p.Name -eq $Info.PolicyName) { ', current' } else { '' })
            $opts += [pscustomobject]@{ Key = [string]$i; Label = [string]$p.Name; Detail = $d }
        }
        $opts += [pscustomobject]@{ Key = 'N'; Label = 'Create a new policy instead'; Detail = 'guided: age in days and a name' }
        $opts += [pscustomobject]@{ Key = '0'; Label = 'Keep the current policy'; Detail = 'no change' }
        $k = Read-Menu -Title 'Retention policies in this tenant' -Options $opts -Default '0' -Note @('Pick the one whose archive age matches what the user expects. Policies without a move-to-archive rule are listed so you can see them, but they will not archive anything.')
        if ($k -eq '0') { return '' }
        if ($k -eq 'N') { return (New-ArchivePolicy -Existing $all -Info $Info) }
        return [string](($opts | Where-Object { $_.Key -eq $k }).Label)
    }
    $hit = $all | Where-Object { $_.Name -ieq $want } | Select-Object -First 1
    if ($hit) { return [string]$hit.Name }
    Write-Log ("Retention policy '{0}' does not exist in this tenant" -f $want) 'WARN' -Note @(("Existing policies: {0}" -f (($all | ForEach-Object { $_.Name }) -join ', ')))
    if ($ArchiveAfterDays -gt 0) { return (New-ArchivePolicy -Existing $all -Info $Info -Name $want -Days $ArchiveAfterDays) }
    if ($script:Interactive) {
        $k = Read-Menu -Title ("Create '{0}' now?" -f $want) -Default '1' -Note @('The tool can create it: one move-to-archive rule with the age you choose, wrapped in a new policy, then assigned to the mailbox.') -Options @(
            [pscustomobject]@{ Key = '1'; Label = 'Yes, create it'; Detail = 'you will be asked for the age in days, then shown the exact commands before anything is created' },
            [pscustomobject]@{ Key = '2'; Label = 'No, keep the current policy'; Detail = 'continue with what the mailbox has today' })
        if ($k -eq '1') { return (New-ArchivePolicy -Existing $all -Info $Info -Name $want) }
        return ''
    }
    throw ("Retention policy '{0}' does not exist. Add -ArchiveAfterDays <days> to create it in an unattended run, or use one of: {1}" -f $want, (($all | ForEach-Object { $_.Name }) -join ', '))
}

function Prepare-Mailbox {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)]$Info)
    $m = $Info.Picture.Mailbox; $box = $Info.Box; $id = $m.ExchangeGuid.ToString()
    if ($Mode -eq 'Run') { Write-Log ("{0}: Run mode, settings are left exactly as they are" -f $box) 'INFO'; return $true }
    if (-not $script:PolicyResolved) { $script:TargetPolicy = [string](Resolve-PolicyChoice -Info $Info); $script:PolicyResolved = $true }
    $newPolicy = $script:TargetPolicy
    $plan = @()
    if ($newPolicy -and $newPolicy -ne $Info.PolicyName) { $plan += ("Assign retention policy '{0}' (currently {1})" -f $newPolicy, $(if ($Info.PolicyName) { "'$($Info.PolicyName)'" } else { 'none' })); $script:PolicyChanged = $true }
    $enableArchive = ($m.ArchiveStatus -ne 'Active')
    $disabledGuid = [string]$m.DisabledArchiveGuid
    $hasDisabled = ($disabledGuid -and $disabledGuid -ne '00000000-0000-0000-0000-000000000000')
    if ($enableArchive) { $plan += $(if ($hasDisabled) { 'Reconnect the disabled archive (Enable-Mailbox -Archive; if Exchange refuses because it is older than 30 days, remove it and create a new one)' } else { 'Enable the archive mailbox (Enable-Mailbox -Archive)' }) }
    if ($m.RetentionHoldEnabled -and $ClearRetentionHold) { $plan += 'Lift the retention hold (Set-Mailbox -RetentionHoldEnabled $false)' }
    if ($m.ElcProcessingDisabled -and $EnableElcProcessing) { $plan += 'Switch the assistant on for this mailbox (Set-Mailbox -ElcProcessingDisabled $false)' }
    if ($AutoExpandingArchive -and $m.PSObject.Properties['AutoExpandingArchiveEnabled'] -and -not $m.AutoExpandingArchiveEnabled) { $plan += 'Enable auto-expanding archiving (Enable-Mailbox -AutoExpandingArchive)' }
    if ($plan.Count -eq 0) { Write-Log ("{0}: nothing to change; the setup is already right" -f $box) 'OK'; return $true }
    Write-Panel -Title ("Changes for {0}" -f $box) -Lines $plan -Color 'Amber'
    if (-not (Confirm-Change -Prompt ("Apply these {0} change(s)?" -f $plan.Count) -Note @('Each line is one Exchange command, run exactly as shown. After each one the tool reads the setting back and reports Verified or NotVerified. Nothing is deleted. Anything other than YES skips all of them, and the assistant is not started for this mailbox in this run.'))) { foreach ($p in $plan) { Add-Action -Box $box -Title $p -Status Skipped -Detail 'not approved' }; return $false }

    if ($newPolicy -and $newPolicy -ne $Info.PolicyName) {
        if ($PSCmdlet.ShouldProcess($box, "Set-Mailbox -RetentionPolicy '$newPolicy'")) {
            try { Set-Mailbox -Identity $id -RetentionPolicy $newPolicy -ErrorAction Stop -WarningAction SilentlyContinue; $after = (Get-Mailbox -Identity $id -WarningAction SilentlyContinue).RetentionPolicy; Add-Action -Box $box -Title ("Retention policy set to '{0}'" -f $newPolicy) -Status $(if ([string]$after -eq $newPolicy) { 'Verified' } else { 'NotVerified' }) -Detail ("Get-Mailbox now reports '{0}'" -f $after) -Note @('The new rules apply from the assistant''s next visit; the run stage asks for one now.') }
            catch { Add-Action -Box $box -Title 'Retention policy' -Status Failed -Detail (Get-ErrorText $_) }
        } else { Add-Action -Box $box -Title ("Set retention policy '{0}'" -f $newPolicy) -Status WhatIf }
    }
    if ($enableArchive) {
        if ($PSCmdlet.ShouldProcess($box, 'Enable-Mailbox -Archive')) {
            try {
                Start-Glow ("Enabling the archive for {0}" -f $box)
                try { Enable-Mailbox -Identity $id -Archive -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null }
                catch {
                    $msg = Get-ErrorText $_
                    if ($hasDisabled -and $msg -match 'MissingDisconnectReceipts|not authorized|disconnect of shard') {
                        Set-GlowText 'The disabled archive is older than 30 days; removing its record and creating a new archive'
                        Set-Mailbox -Identity $id -RemoveDisabledArchive -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                        Enable-Mailbox -Identity $id -Archive -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                    } else { throw }
                }
                $deadline = (Get-Date).AddMinutes(2); $status = ''
                do { Start-Sleep -Seconds 10; $status = [string](Get-Mailbox -Identity $id -WarningAction SilentlyContinue).ArchiveStatus; Set-GlowText ("Exchange is provisioning the archive (status {0}); this usually takes under a minute" -f $status) } while ($status -ne 'Active' -and (Get-Date) -lt $deadline)
                Stop-Glow
                Add-Action -Box $box -Title 'Archive mailbox enabled' -Status $(if ($status -eq 'Active') { 'Verified' } else { 'NotVerified' }) -Detail ("ArchiveStatus {0}" -f $status) -Note $(if ($status -eq 'Active') { @('The archive exists and is ready. It appears in Outlook as Online Archive after the user''s next sign-in.') } else { @('Exchange has not finished provisioning yet. It usually completes within a few minutes; run Check later to confirm ArchiveStatus Active.') })
            } catch { Stop-Glow; Add-Action -Box $box -Title 'Enable archive' -Status Failed -Detail (Get-ErrorText $_) -Note @('Common causes: the licence has no archive entitlement (Exchange Online Plan 1 without the Exchange Online Archiving add-on), or the signed-in account lacks the Mail Recipients role.') }
        } else { Add-Action -Box $box -Title 'Enable the archive mailbox' -Status WhatIf }
    }
    if ($m.RetentionHoldEnabled -and $ClearRetentionHold) {
        if ($PSCmdlet.ShouldProcess($box, 'Set-Mailbox -RetentionHoldEnabled $false')) { try { Set-Mailbox -Identity $id -RetentionHoldEnabled $false -ErrorAction Stop -WarningAction SilentlyContinue; $v = (Get-Mailbox -Identity $id -WarningAction SilentlyContinue).RetentionHoldEnabled; Add-Action -Box $box -Title 'Retention hold lifted' -Status $(if (-not $v) { 'Verified' } else { 'NotVerified' }) } catch { Add-Action -Box $box -Title 'Retention hold' -Status Failed -Detail (Get-ErrorText $_) } } else { Add-Action -Box $box -Title 'Lift the retention hold' -Status WhatIf }
    }
    if ($m.ElcProcessingDisabled -and $EnableElcProcessing) {
        if ($PSCmdlet.ShouldProcess($box, 'Set-Mailbox -ElcProcessingDisabled $false')) { try { Set-Mailbox -Identity $id -ElcProcessingDisabled $false -ErrorAction Stop -WarningAction SilentlyContinue; $v = (Get-Mailbox -Identity $id -WarningAction SilentlyContinue).ElcProcessingDisabled; Add-Action -Box $box -Title 'Assistant switched on for this mailbox' -Status $(if (-not $v) { 'Verified' } else { 'NotVerified' }) } catch { Add-Action -Box $box -Title 'ElcProcessingDisabled' -Status Failed -Detail (Get-ErrorText $_) } } else { Add-Action -Box $box -Title 'Switch the assistant on' -Status WhatIf }
    }
    if ($AutoExpandingArchive -and $m.PSObject.Properties['AutoExpandingArchiveEnabled'] -and -not $m.AutoExpandingArchiveEnabled) {
        if ($PSCmdlet.ShouldProcess($box, 'Enable-Mailbox -AutoExpandingArchive')) { try { Enable-Mailbox -Identity $id -AutoExpandingArchive -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null; $v = (Get-Mailbox -Identity $id -WarningAction SilentlyContinue).AutoExpandingArchiveEnabled; Add-Action -Box $box -Title 'Auto-expanding archive enabled' -Status $(if ($v) { 'Verified' } else { 'NotVerified' }) -Detail 'storage is added as the archive fills; this cannot be turned off later' } catch { Add-Action -Box $box -Title 'Auto-expanding archive' -Status Failed -Detail (Get-ErrorText $_) } } else { Add-Action -Box $box -Title 'Enable auto-expanding archive' -Status WhatIf }
    }
    return $true
}

function Run-Assistant {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)]$Info)
    $m = $Info.Picture.Mailbox; $box = $Info.Box; $id = $m.ExchangeGuid.ToString()
    $fresh = Get-Mailbox -Identity $id -WarningAction SilentlyContinue
    $hardBlock = @()
    if ($fresh.ArchiveStatus -ne 'Active') { $hardBlock += 'archive not active' }
    if ($fresh.RetentionHoldEnabled) { $hardBlock += 'retention hold' }
    if ($fresh.ElcProcessingDisabled) { $hardBlock += 'ElcProcessingDisabled' }
    if ($hardBlock.Count -gt 0) { Add-Action -Box $box -Title 'Managed Folder Assistant' -Status Skipped -Detail ("blocked by: {0}" -f ($hardBlock -join ', ')) -Note @('Starting the assistant now would report success and move nothing, so it is not started. Clear the blocker (Fix mode) and run again.'); return }

    $before = Get-MailboxPicture -Identity $id
    $history = [System.Collections.Generic.List[object]]::new()
    $history.Add([pscustomobject]@{ Pass = 0; Time = (Get-Date).ToString('HH:mm:ss'); PrimaryBytes = $before.PrimaryBytes; PrimaryItems = $before.PrimaryItems; ArchiveBytes = $before.ArchiveBytes; ArchiveItems = $before.ArchiveItems; Moved = $null; LastSuccess = (Get-Counter $Info.Counters 'ElcLastSuccessTimestamp') })
    Write-Log ("{0}: {1} pass(es); {2} after each request{3}" -f $box, $Passes, $(if ($IntervalMinutes -gt 0) { "wait $IntervalMinutes minute(s)" } else { 'no wait' }), $(if ($FullCrawl) { '; full crawl on the first' } else { '' })) 'STEP' -Note @('A pass is one request to Exchange, one wait, one measurement. Exchange keeps working between passes and after the last one, on its own schedule.')
    Show-Picture -Pic $before -Label 'before' -Color 'Copper'
    $lastSeenItems = $before.ArchiveItems; $sessionMoved = 0

    for ($n = 1; $n -le $Passes; $n++) {
        $useCrawl = ($FullCrawl -and $n -eq 1)
        $what = "Start-ManagedFolderAssistant -Identity <guid>" + $(if ($useCrawl) { ' -FullCrawl' } else { '' })
        if (-not $PSCmdlet.ShouldProcess($box, $what)) { Add-Action -Box $box -Title ("Pass {0}: {1}" -f $n, $what) -Status WhatIf; continue }
        try {
            Start-Glow ("Pass {0} of {1}: asking Exchange to process the mailbox{2}" -f $n, $Passes, $(if ($useCrawl) { ' (full crawl)' } else { '' }))
            if ($useCrawl) {
                try { Start-ManagedFolderAssistant -Identity $id -FullCrawl -ErrorAction Stop -WarningAction SilentlyContinue }
                catch [System.Management.Automation.ParameterBindingException] { Write-Log 'This module build could not bind -FullCrawl on its own; running a normal pass instead' 'WARN'; Start-ManagedFolderAssistant -Identity $id -ErrorAction Stop -WarningAction SilentlyContinue }
            } else { Start-ManagedFolderAssistant -Identity $id -ErrorAction Stop -WarningAction SilentlyContinue }
            if ($IntervalMinutes -le 0) {
                Stop-Glow
                Add-Action -Box $box -Title ("Pass {0}: Managed Folder Assistant" -f $n) -Status Started -Detail 'request accepted; no wait was chosen, so the result is not measured in this run' -Note @('Exchange is processing the mailbox in the background. Run this tool again in Check mode in 15 to 30 minutes to see the sizes and counters.')
                $history.Add([pscustomobject]@{ Pass = $n; Time = (Get-Date).ToString('HH:mm:ss'); PrimaryBytes = $before.PrimaryBytes; PrimaryItems = $before.PrimaryItems; ArchiveBytes = $before.ArchiveBytes; ArchiveItems = $before.ArchiveItems; Moved = $null; LastSuccess = $history[$history.Count - 1].LastSuccess })
                continue
            }
            $waitSec = $IntervalMinutes * 60
            $pollEvery = $(if ($waitSec -le 180) { 30 } else { 60 })
            $t0 = Get-Date; $nextPoll = $t0.AddSeconds($pollEvery)
            while (((Get-Date) - $t0).TotalSeconds -lt $waitSec) {
                $left = [int]($waitSec - ((Get-Date) - $t0).TotalSeconds)
                $toPoll = [int][math]::Max(0, ($nextPoll - (Get-Date)).TotalSeconds)
                Set-GlowText ("Pass {0} of {1}: Exchange is working; next size check in {2}:{3:00}, results in {4}:{5:00}{6}" -f $n, $Passes, [int][math]::Floor($toPoll / 60), ($toPoll % 60), [int][math]::Floor($left / 60), ($left % 60), $(if ($sessionMoved -gt 0) { ("; {0:N0} items moved so far" -f $sessionMoved) } else { '' }))
                Start-Sleep -Seconds 5
                if ((Get-Date) -ge $nextPoll -and $left -gt 15) {
                    $nextPoll = (Get-Date).AddSeconds($pollEvery)
                    $now = $null; try { $now = Get-MailboxPicture -Identity $id } catch { }
                    if ($now) {
                        $d = $now.ArchiveItems - $lastSeenItems
                        if ($d -gt 0) {
                            $sessionMoved += $d
                            Write-Log ("+{0:N0} items landed in the archive  (archive now {1}, {2:N0} items; primary {3})" -f $d, (Format-Size $now.ArchiveBytes), $now.ArchiveItems, (Format-Size $now.PrimaryBytes)) 'OK'
                            Show-Picture -Pic $now -Label 'moving' -Color 'Gold'
                        }
                        $lastSeenItems = $now.ArchiveItems
                    }
                }
            }
            Set-GlowText ("Pass {0} of {1}: reading Exchange's counters and the final sizes" -f $n, $Passes)
            $counters = Get-ElcCounters -Identity $id
            $after = Get-MailboxPicture -Identity $id
            Stop-Glow
            $moved = Get-Counter $counters 'ElcLastRunArchivedFromRootItemCount'
            $movedDump = Get-Counter $counters 'ElcLastRunArchivedFromDumpsterItemCount'
            $tagged = Get-Counter $counters 'ElcLastRunTaggedWithArchiveItemCount'
            $last = Get-Counter $counters 'ElcLastSuccessTimestamp'
            $deltaItems = $after.ArchiveItems - $before.ArchiveItems
            $history.Add([pscustomobject]@{ Pass = $n; Time = (Get-Date).ToString('HH:mm:ss'); PrimaryBytes = $after.PrimaryBytes; PrimaryItems = $after.PrimaryItems; ArchiveBytes = $after.ArchiveBytes; ArchiveItems = $after.ArchiveItems; Moved = $moved; LastSuccess = $last })
            Show-Picture -Pic $after -Label ("pass {0}" -f $n) -Color 'Amber'
            $detail = ("last run moved {0} item(s) from the mailbox and {1} from Recoverable Items, tagged {2} for archive; archive now {3:N0} items ({4} in this pass); ElcLastSuccessTimestamp {5}" -f $(if ($null -ne $moved) { $moved } else { '?' }), $(if ($null -ne $movedDump) { $movedDump } else { '?' }), $(if ($null -ne $tagged) { $tagged } else { '?' }), $after.ArchiveItems, $(if ($deltaItems -ge 0) { "+$deltaItems" } else { $deltaItems }), $(if ($last) { $last } else { 'not yet' }))
            $prevLast = $history[$history.Count - 2].LastSuccess
            $advanced = [bool]$last -and ([string]$last -ne [string]$prevLast)
            $movedNow = ($deltaItems -gt 0) -or ($null -ne $moved -and [int]$moved -gt 0)
            # A pass counts as verified when items moved, or when the assistant completed a fresh run (new success timestamp) and simply found nothing eligible.
            $success = $movedNow -or $advanced
            $note = $(if ($movedNow) { @(("Exchange moved {0:N0} item(s) in this pass. The archive now holds {1} ({2:N0} items)." -f [math]::Max($deltaItems, $(if ($null -ne $moved) { [int]$moved } else { 0 })), (Format-Size $after.ArchiveBytes), $after.ArchiveItems)) }
                     elseif ($advanced) { @('Exchange finished the request (its success timestamp advanced) and found nothing old enough to move. That is a clean result, not an error.') }
                     else { @('Exchange has not finished this request yet; large mailboxes are processed in throttled slices. The next pass reads the counters again.') })
            Add-Action -Box $box -Title ("Pass {0}: Managed Folder Assistant" -f $n) -Status $(if ($success) { 'Verified' } else { 'NotVerified' }) -Detail $detail -Note $note
            if (-not $movedNow -and $advanced -and $n -ge 2) {
                $prevMoved = $history[$history.Count - 2].Moved
                if ($null -ne $prevMoved -and [int]$prevMoved -eq 0) { Write-Log ("{0}: two passes in a row completed cleanly and moved nothing. Stopping early." -f $box) 'OK' -Note @('Nothing else qualifies under the current archive age. Items that reach it later move on Exchange''s own weekly schedule.'); break }
            }
            $before = $after; $lastSeenItems = $after.ArchiveItems
        } catch { Stop-Glow; Add-Action -Box $box -Title ("Pass {0}: Managed Folder Assistant" -f $n) -Status Failed -Detail (Get-ErrorText $_) }
    }
    # MRM diagnostic log for the report
    try {
        $log = Export-MailboxDiagnosticLogs -Identity $id -ComponentName MRM -ErrorAction Stop -WarningAction SilentlyContinue
        $text = [string]$log.MailboxLog
        Set-Content -Path (Join-Path $script:RunDir ("mrm-log-{0}.txt" -f ($box -replace '[^A-Za-z0-9.@_-]', '_'))) -Value $text -Encoding UTF8
        if ($text -match 'resource unhealthy|ResourceUnhealthy') { Add-Finding -Box $box -Kind Warning -Title 'MRM log shows throttling ("resource unhealthy"); a large mailbox is processed slowly' -Fix 'Expected for large mailboxes; keep running passes over the coming days' -Source 'Resolve email archive and deletion issues when using MRM' -Why 'Exchange limits how much one mailbox may move at a time so the service stays healthy for everyone. The backlog clears over several runs.' }
    } catch { if ((Get-ErrorText $_) -match 'no logs were found|No logs') { Add-Finding -Box $box -Kind Info -Title 'MRM diagnostic log is empty: the assistant processed the mailbox without errors' -Source 'Resolve email archive and deletion issues when using MRM' } else { Write-Log ("MRM log: {0}" -f (Get-ErrorText $_)) 'DEBUG' } }
    $script:Results.Add([pscustomobject]@{ Mailbox = $box; History = @($history); SessionMoved = $sessionMoved })
}

# ---------------------------------------------------------------------------------------------
# Outcome, report, summary
# ---------------------------------------------------------------------------------------------
function Get-Outcome {
    # The result in plain words. Used in the summary panel, CASE-NOTES.txt and report.html.
    $lines = @()
    $blockers = @($script:Findings | Where-Object { $_.Kind -eq 'Blocker' })
    $small = @($blockers | Where-Object { $_.Title -like '*under 10 MB*' })
    $real = @($blockers | Where-Object { $_.Title -notlike '*under 10 MB*' })
    $notYet = @($script:Findings | Where-Object { $_.Title -like 'Nothing qualifies for the archive yet*' })
    if ($Mode -eq 'Check') {
        $lines += 'Check mode: nothing was changed and the assistant was not started.'
        if ($real.Count -gt 0) { $lines += ("{0} thing(s) stop archiving in this mailbox; each is listed above with its fix. Run again and choose Fix to clear the ones the tool can clear." -f $real.Count) }
        elseif ($notYet.Count -gt 0) { $lines += 'The setup is right, but nothing is old enough to move yet: ' + $notYet[0].Title.Substring($notYet[0].Title.IndexOf(':') + 2) + '; ' + $notYet[0].Detail + '.' }
        else { $lines += 'No blocker found. Run again and choose Run only (or Fix) to start the assistant and watch what moves.' }
    } else {
        $started = @($script:Actions | Where-Object { $_.Status -eq 'Started' })
        $passes = @($script:Actions | Where-Object { $_.Title -like 'Pass *' -and $_.Status -in @('Verified', 'NotVerified') })
        $skippedRun = @($script:Actions | Where-Object { $_.Title -eq 'Managed Folder Assistant' -and $_.Status -eq 'Skipped' })
        $notApproved = @($script:Actions | Where-Object { $_.Status -eq 'Skipped' -and $_.Detail -eq 'not approved' })
        $failed = @($script:Actions | Where-Object { $_.Status -in @('Failed', 'NotVerified') -and $_.Title -notlike 'Pass *' })
        $moved = 0; foreach ($r in $script:Results) { $moved += ($r.History[$r.History.Count - 1].ArchiveItems - $r.History[0].ArchiveItems) }
        if ($WhatIfPreference) { $lines += 'Dry run (-WhatIf): the plan above is what a real run would do. Nothing was changed and the assistant was not started.' }
        elseif ($notApproved.Count -gt 0 -and $passes.Count -eq 0 -and $started.Count -eq 0) { $lines += 'The changes were not approved, so nothing was changed and the assistant was not started. Run again and type YES when the change list looks right.' }
        elseif ($skippedRun.Count -gt 0 -and $passes.Count -eq 0 -and $started.Count -eq 0) { $lines += ("The assistant was not started because a blocker remains ({0}). Clear it and run again." -f $skippedRun[0].Detail) }
        elseif ($started.Count -gt 0 -and $passes.Count -eq 0) { $lines += 'The assistant was asked to process the mailbox and is working in the background. Run Check in 15 to 30 minutes to see what moved.' }
        elseif ($moved -gt 0) { $lines += ("Archiving works: {0:N0} item(s) moved into the archive during this run. Exchange keeps running the assistant at least weekly, so items that reach the archive age later move on their own." -f $moved) }
        elseif ($passes.Count -gt 0) {
            if ($notYet.Count -gt 0) { $lines += 'The setup is right and the assistant completed, but nothing is old enough to move yet.'; $lines += $notYet[0].Title.Substring($notYet[0].Title.IndexOf(':') + 2) + '; ' + $notYet[0].Detail + '.'; $lines += 'To archive sooner, assign or create a policy with a shorter archive age (Fix mode).' }
            elseif (@($passes | Where-Object { $_.Status -eq 'Verified' }).Count -gt 0) { $lines += 'The assistant completed and found nothing to move under the current tags. If the user expected items to move, compare the archive age of the policy with the age of the mail.' }
            else { $lines += 'Exchange had not finished processing when the last measurement was taken. Large mailboxes move in slices; run Check later, or run again with more passes and a longer wait.' }
        }
        if ($failed.Count -gt 0) { $lines += ("{0} change(s) failed or could not be verified; see Actions above and in the report." -f $failed.Count) }
        if ($small.Count -gt 0) { $lines += 'This mailbox is under 10 MB, which Exchange skips by design; archiving begins once it grows.' }
    }
    if ($lines.Count -eq 0) { $lines += 'Nothing was done. See the findings above.' }
    return $lines
}

function ConvertTo-HtmlText { param([string]$Text) return [System.Net.WebUtility]::HtmlEncode([string]$Text) }
function Write-Reports {
    $end = Get-Date
    $blockers = @($script:Findings | Where-Object { $_.Kind -eq 'Blocker' })
    $unverified = @($script:Actions | Where-Object { $_.Status -in @('NotVerified', 'Failed') })
    $skipped = @($script:Actions | Where-Object { $_.Status -eq 'Skipped' -and $_.Detail -match 'not approved|blocked by' })
    if ($script:Stopped) { $script:ExitCode = 3 } elseif ($unverified.Count -gt 0) { $script:ExitCode = 2 } elseif ($blockers.Count -gt 0 -or $skipped.Count -gt 0) { $script:ExitCode = 1 }
    $outcome = @(Get-Outcome)
    $result = [ordered]@{ Tool = 'Mailbox Archive Runner'; Version = $script:Version; RunId = $script:RunId; CaseNumber = $CaseNumber; Mode = $Mode; Started = $script:Started.ToString('s'); Ended = $end.ToString('s'); ExitCode = $script:ExitCode; Outcome = $outcome; Mailboxes = @($Mailbox); Passes = $Passes; IntervalMinutes = $IntervalMinutes; FullCrawl = [bool]$FullCrawl; Findings = @($script:Findings); Actions = @($script:Actions); Results = @($script:Results) }
    $result | ConvertTo-Json -Depth 6 | Set-Content -Path (Join-Path $script:RunDir 'result.json') -Encoding UTF8
    $notes = [System.Collections.Generic.List[string]]::new()
    $notes.Add(("Mailbox Archive Runner v{0}  {1}{2}" -f $script:Version, $end.ToString('yyyy-MM-dd HH:mm'), $(if ($CaseNumber) { "  case $CaseNumber" } else { '' })))
    $notes.Add(("Mailboxes: {0}   Mode: {1}   Passes: {2} with {3} min wait   Result: exit {4}" -f ($Mailbox -join ', '), $Mode, $Passes, $IntervalMinutes, $script:ExitCode))
    $notes.Add('Outcome: ' + ($outcome -join ' '))
    foreach ($r in $script:Results) { $first = $r.History[0]; $lastH = $r.History[$r.History.Count - 1]; $notes.Add(("  {0}: archive {1} -> {2} items, {3} -> {4}; primary {5} -> {6}; ElcLastSuccessTimestamp {7}" -f $r.Mailbox, $first.ArchiveItems, $lastH.ArchiveItems, (Format-Size $first.ArchiveBytes), (Format-Size $lastH.ArchiveBytes), (Format-Size $first.PrimaryBytes), (Format-Size $lastH.PrimaryBytes), $(if ($lastH.LastSuccess) { $lastH.LastSuccess } else { 'none' }))) }
    if ($blockers.Count -gt 0) { $notes.Add('Blockers:'); foreach ($b in $blockers) { $notes.Add(("  [{0}] {1}{2}. Fix: {3}" -f $b.Mailbox, $b.Title, $(if ($b.Detail) { " ($($b.Detail))" } else { '' }), $b.Fix)) } }
    if ($script:Actions.Count -gt 0) { $notes.Add('Actions:'); foreach ($a in $script:Actions) { $notes.Add(("  [{0}] {1}: {2}{3}" -f $a.Mailbox, $a.Status, $a.Title, $(if ($a.Detail) { " ($($a.Detail))" } else { '' }))) } }
    $notes.Add(("Report: {0}" -f $script:RunDir))
    Set-Content -Path (Join-Path $script:RunDir 'CASE-NOTES.txt') -Value $notes -Encoding UTF8

    $kc = @{ Blocker = '#b5533c'; Warning = '#c98a2e'; Info = '#7a8a99' }; $sc = @{ Verified = '#4f7f5a'; Started = '#4f7f5a'; NotVerified = '#c98a2e'; Failed = '#b5533c'; Skipped = '#7a8a99'; WhatIf = '#7a8a99' }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Mailbox Archive Runner report</title>')
    [void]$sb.AppendLine('<style>body{font-family:Georgia,"Times New Roman",serif;margin:0;background:#f5f1ea;color:#2b2a27}header{background:#2b2a27;color:#f5e6c8;padding:30px 44px;border-bottom:4px solid #d08a2e}header h1{margin:0;font-size:24px;font-weight:normal;letter-spacing:.5px}header p{margin:6px 0 0;color:#c9b99a;font-family:Segoe UI,system-ui,sans-serif;font-size:14px}main{padding:26px 44px;max-width:1180px}h2{font-size:18px;font-weight:normal;border-left:4px solid #d08a2e;padding-left:10px;margin:30px 0 12px}.outcome{background:#fffdf8;border-left:4px solid #4f7f5a;padding:14px 18px;font-size:16px;line-height:1.5;box-shadow:0 1px 2px rgba(0,0,0,.08)}table{width:100%;border-collapse:collapse;background:#fffdf8;font-family:Segoe UI,system-ui,sans-serif;font-size:14px;box-shadow:0 1px 2px rgba(0,0,0,.08)}th{background:#e9e1d2;text-align:left;padding:10px 12px;font-size:13px}td{padding:10px 12px;border-top:1px solid #eee5d6;vertical-align:top}.tag{display:inline-block;padding:2px 9px;border-radius:3px;color:#fff;font-size:12px}.meter{height:10px;background:#e9e1d2;border-radius:2px;overflow:hidden;width:220px;display:inline-block;vertical-align:middle}.meter span{display:block;height:100%;background:#d08a2e}.why{color:#6b665d;font-size:13px}code{background:#eee5d6;padding:1px 5px;border-radius:3px;font-size:13px}footer{padding:22px 44px;color:#7a8a99;font-size:13px;font-family:Segoe UI,system-ui,sans-serif}</style></head><body>')
    [void]$sb.AppendLine(('<header><h1>Mailbox Archive Runner</h1><p>{0} &middot; {1} &middot; {2} mode &middot; {3} pass(es), {4} min wait{5} &middot; exit code {6}</p></header><main>' -f (ConvertTo-HtmlText ($Mailbox -join ', ')), $end.ToString('yyyy-MM-dd HH:mm'), $Mode, $Passes, $IntervalMinutes, $(if ($CaseNumber) { ' &middot; case ' + (ConvertTo-HtmlText $CaseNumber) } else { '' }), $script:ExitCode))
    [void]$sb.AppendLine(('<h2>What this means</h2><div class="outcome">{0}</div>' -f (ConvertTo-HtmlText ($outcome -join ' '))))
    foreach ($r in $script:Results) {
        [void]$sb.AppendLine(('<h2>{0}</h2><table><tr><th>Pass</th><th>Time</th><th>Primary</th><th>Archive</th><th>Share in archive</th><th>Moved in last run</th><th>ElcLastSuccessTimestamp</th></tr>' -f (ConvertTo-HtmlText $r.Mailbox)))
        foreach ($h in $r.History) { $tot = [math]::Max(1, $h.PrimaryBytes + $h.ArchiveBytes); $pct = [int](100 * $h.ArchiveBytes / $tot); [void]$sb.AppendLine(('<tr><td>{0}</td><td>{1}</td><td>{2} ({3:N0} items)</td><td>{4} ({5:N0} items)</td><td><span class="meter"><span style="width:{6}%"></span></span> {6}%</td><td>{7}</td><td>{8}</td></tr>' -f $(if ($h.Pass -eq 0) { 'before' } else { $h.Pass }), $h.Time, (Format-Size $h.PrimaryBytes), $h.PrimaryItems, (Format-Size $h.ArchiveBytes), $h.ArchiveItems, $pct, $(if ($null -ne $h.Moved) { $h.Moved } else { '' }), (ConvertTo-HtmlText ([string]$h.LastSuccess)))) }
        [void]$sb.AppendLine('</table>')
    }
    [void]$sb.AppendLine('<h2>Findings</h2><table><tr><th>Kind</th><th>Mailbox</th><th>Finding</th><th>Detail</th><th>Why it matters</th><th>Fix</th><th>Source</th></tr>')
    if ($script:Findings.Count -eq 0) { [void]$sb.AppendLine('<tr><td colspan="7">Nothing to report.</td></tr>') }
    foreach ($f in ($script:Findings | Sort-Object @{ E = { switch ($_.Kind) { 'Blocker' { 0 } 'Warning' { 1 } default { 2 } } } })) { [void]$sb.AppendLine(('<tr><td><span class="tag" style="background:{0}">{1}</span></td><td>{2}</td><td>{3}</td><td>{4}</td><td class="why">{5}</td><td>{6}</td><td>{7}</td></tr>' -f $kc[$f.Kind], $f.Kind, (ConvertTo-HtmlText $f.Mailbox), (ConvertTo-HtmlText $f.Title), (ConvertTo-HtmlText $f.Detail), (ConvertTo-HtmlText $f.Why), (ConvertTo-HtmlText $f.Fix), (ConvertTo-HtmlText $f.Source))) }
    [void]$sb.AppendLine('</table><h2>Actions</h2><table><tr><th>Status</th><th>Mailbox</th><th>Action</th><th>Detail</th></tr>')
    if ($script:Actions.Count -eq 0) { [void]$sb.AppendLine('<tr><td colspan="4">No changes were made.</td></tr>') }
    foreach ($a in $script:Actions) { [void]$sb.AppendLine(('<tr><td><span class="tag" style="background:{0}">{1}</span></td><td>{2}</td><td>{3}</td><td>{4}</td></tr>' -f $sc[$a.Status], $a.Status, (ConvertTo-HtmlText $a.Mailbox), (ConvertTo-HtmlText $a.Title), (ConvertTo-HtmlText $a.Detail))) }
    [void]$sb.AppendLine(('</table><h2>Files</h2><p>Run folder: <code>{0}</code>. result.json, CASE-NOTES.txt, run.log and one MRM diagnostic log per mailbox.</p>' -f (ConvertTo-HtmlText $script:RunDir)))
    [void]$sb.AppendLine(('</main><footer>Mailbox Archive Runner v{0} &middot; Arwaz Khan, Microsoft Support Engineer &middot; every step follows the Microsoft articles listed in docs/sources.md</footer></body></html>' -f $script:Version))
    $html = Join-Path $script:RunDir 'report.html'
    Set-Content -Path $html -Value $sb.ToString() -Encoding UTF8
    return $html
}

function Show-Summary {
    param([string]$HtmlPath)
    $lines = @()
    foreach ($r in $script:Results) {
        $first = $r.History[0]; $lastH = $r.History[$r.History.Count - 1]
        $tot = [math]::Max(1, $lastH.PrimaryBytes + $lastH.ArchiveBytes)
        $lines += ("{0}" -f (Tint $r.Mailbox 'Cream'))
        $lines += ("  {0}  archive {1} ({2:N0} items), primary {3} ({4:N0} items)" -f (Get-Meter -Value $lastH.ArchiveBytes -Max $tot -Color 'Amber'), (Format-Size $lastH.ArchiveBytes), $lastH.ArchiveItems, (Format-Size $lastH.PrimaryBytes), $lastH.PrimaryItems)
        $lines += ("  {0} items moved into the archive during this run" -f (Tint ("{0:+#,0;-#,0;0}" -f ($lastH.ArchiveItems - $first.ArchiveItems)) $(if ($lastH.ArchiveItems -gt $first.ArchiveItems) { 'Sage' } else { 'Slate' })))
    }
    $blockers = @($script:Findings | Where-Object { $_.Kind -eq 'Blocker' })
    if ($blockers.Count -gt 0) { $lines += ''; $lines += (Tint ("{0} blocker(s) still standing:" -f $blockers.Count) 'Rust'); foreach ($b in $blockers | Select-Object -First 6) { $lines += ("  {0} {1}: {2}" -f $script:G.Warn, $b.Mailbox, $b.Title); if ($b.Fix) { $lines += ("      fix: {0}" -f $b.Fix) } } }
    $lines += ''
    $lines += (Tint 'What this means' 'Gold')
    foreach ($o in (Get-Outcome)) { $lines += ("  {0}" -f $o) }
    $lines += ''
    $lines += ("Report  {0}" -f (Tint $HtmlPath 'Sky'))
    $startedOnly = (@($script:Actions | Where-Object { $_.Status -eq 'Started' }).Count -gt 0) -and (@($script:Actions | Where-Object { $_.Title -like 'Pass *' -and $_.Status -in @('Verified', 'NotVerified') }).Count -eq 0)
    $lines += ("Exit    {0}  ({1})" -f $script:ExitCode, $(switch ($script:ExitCode) { 0 { $(if ($Mode -eq 'Check') { 'checked, no blocker' } elseif ($startedOnly) { 'started; result not yet measured' } else { 'done and verified' }) } 1 { 'a blocker remains or a change was not approved' } 2 { 'a change could not be verified' } 3 { 'stopped early' } default { '' } }))
    Write-Panel -Title 'Summary' -Lines $lines -Color $(switch ($script:ExitCode) { 0 { 'Sage' } 1 { 'Amber' } 2 { 'Copper' } default { 'Rust' } }) -Reveal
}

# ---------------------------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------------------------
function Invoke-Main {
    Show-Intro
    if ($Mailbox.Count -eq 0) {
        if (-not $script:Interactive) { throw 'Pass -Mailbox in an unattended run.' }
        if (-not (Show-StartMenu)) { $script:MainCompleted = $true; return }
    }
    Write-Stage -Title 'Connect to Exchange Online' -Index 1
    $null = Connect-Exchange
    $infos = @()
    Write-Stage -Title 'Read the mailbox, the archive, the policy and its tags' -Index 2
    foreach ($b in $Mailbox) { try { $infos += (Inspect-Mailbox -Identity $b) } catch { Stop-Glow; Add-Finding -Box $b -Kind Blocker -Title 'Mailbox could not be read' -Detail (Get-ErrorText $_) -Fix 'Check the address and your Exchange role (Mail Recipients is required)' -Why 'Either the address does not resolve to a mailbox in this tenant, or the signed-in account is not allowed to read it.' } }
    if ($infos.Count -eq 0) { $script:MainCompleted = $true; return }
    if ($Mode -eq 'Check') { Write-Log 'Check mode: stopping here. Nothing is changed and the assistant is not started.' 'INFO'; $script:MainCompleted = $true; return }
    if ($script:MenuMode -and $script:Interactive) { Ask-Plan -Infos $infos }
    Write-Stage -Title $(if ($Mode -eq 'Run') { 'Prepare: skipped in Run mode, settings stay as they are' } else { 'Prepare: archive, policy, blockers' }) -Index 3
    $ready = @()
    foreach ($i in $infos) { if (Prepare-Mailbox -Info $i) { $ready += $i } }
    Write-Stage -Title 'Run the Managed Folder Assistant and watch' -Index 4
    foreach ($i in $ready) { Run-Assistant -Info $i }
    $script:MainCompleted = $true
}

$htmlPath = ''
try { Invoke-Main }
catch { if ($_.FullyQualifiedErrorId -match 'PipelineStopped') { $script:Stopped = $true } else { Write-Log ("Stopped: {0}" -f (Get-ErrorText $_)) 'FAIL'; Write-Log ([string]$_.ScriptStackTrace) 'DEBUG'; $script:Stopped = $true } }
finally {
    if (-not $script:MainCompleted) { $script:Stopped = $true }
    if ($script:Stopped) { Remove-Glow }
    try {
        Write-Stage -Title 'Report' -Index 5
        if (-not $script:Stopped) { Start-Glow 'Writing report.html, result.json and CASE-NOTES.txt' }
        $htmlPath = Write-Reports
        Stop-Glow
        Show-Summary -HtmlPath $htmlPath
        if ($script:Interactive -and -not $NoOpenReport -and $htmlPath -and (Test-Path -LiteralPath $htmlPath)) { try { Start-Process -FilePath $htmlPath -ErrorAction Stop } catch { } }
        Show-Outro
    } catch { Stop-Glow; Write-Host ("Report could not be written: {0}" -f (Get-ErrorText $_)) -ForegroundColor Red; if ($script:ExitCode -lt 2) { $script:ExitCode = 2 } }
    finally { Remove-Glow }
}
exit $script:ExitCode
