# Troubleshooting the tool

## "Unexpected token" errors with characters like â•” or â–°

The script was read in the wrong encoding. It ships as UTF-8 with a byte-order mark so both PowerShell editions decode the symbols; if the mark was stripped by a copy-paste, download the file again or run it with PowerShell 7 (`pwsh -File .\Invoke-MailboxArchive.ps1`).

## The module will not install

`Install-Module ExchangeOnlineManagement -Scope CurrentUser` needs the PowerShell Gallery to be reachable and, on Windows PowerShell 5.1, a current NuGet provider. Run `Install-PackageProvider NuGet -Force -Scope CurrentUser` first if the Gallery step fails, or install the module once from a machine that can reach it.

## "Mailbox could not be read"

The identity did not resolve, or your account lacks the Mail Recipients role. Use the UPN, and check your role group in the Exchange admin center.

## Every pass says NotVerified and nothing moves

Read the findings first: a blocker (hold, ELC disabled, no archive tag, archive not active) means the assistant did nothing by design. If there is no blocker and the MRM log mentions "resource unhealthy", the mailbox is large and is being throttled; keep running passes over the coming days. If the policy has only personal tags, only the folders the user tagged will move.

## "Create retention policy" failed

The message usually names a missing role. Creating tags and policies needs Retention Management, which the Organization Management, Compliance Management and Records Management role groups include; Recipient Management alone is enough to assign a policy but not to create one. The mailbox run continues with the current policy.

## Every pass says "found nothing old enough to move"

Read the warning printed at the read stage: it names the age of the oldest mail item and the archive age of the policy, and the date the first item will qualify. That is the policy working as written. To archive sooner, choose Fix and create a policy with a shorter age, then run with a full crawl.

## "Parameter set cannot be resolved" on -FullCrawl

The tool catches this and runs a normal pass instead, with a warning in the log. Update the ExchangeOnlineManagement module if you want the full crawl.

## Where are the files

The run folder path is in the summary panel and in CASE-NOTES.txt. Default `%LOCALAPPDATA%\MailboxArchiveRunner\<run id>`.
