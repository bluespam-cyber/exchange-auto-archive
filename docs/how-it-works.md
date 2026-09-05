# How it works

## The one thing to understand

Messaging records management in Exchange Online is asynchronous. You assign a policy; the Managed Folder Assistant applies it when it gets to the mailbox, which Microsoft says is at least once every seven days. It stamps items with tags, then moves the ones past their age to the archive. If a precondition is missing, it does nothing and says nothing.

So the tool treats "run the assistant" as the last step, not the first.

## Three modes

- **Check** reads everything and reports. Nothing is changed and the assistant is not started. Use it first on an unfamiliar mailbox, and later to read the result of a "start and leave" run.
- **Fix** (default) reads, then prepares with your approval, then runs and measures.
- **Run** reads, changes no settings, runs and measures. For a mailbox whose setup is already right.

## The guided start

With no parameters the tool asks only two things before connecting: the mode and the mailbox. Everything else is asked after the mailbox has been read, so the questions are about real state: the current policy is named with its archive age; the tenant list shows each policy's archive age; a hold or an ELC block is asked about only when it is actually set. Each question carries a short explanation of what it is for and what every answer does. Enter keeps the suggested answer. When parameters are given on the command line the matching questions are skipped.

## The five stages

1. **Connect.** Reuses an open Exchange Online session; otherwise Connect-ExchangeOnline. Installs the ExchangeOnlineManagement module with your approval if it is absent.
2. **Read.** Get-Mailbox, Get-MailboxStatistics (primary and archive), Get-RetentionPolicy and Get-RetentionPolicyTag for the assigned policy, Get-User for the account state, Get-MailboxFolderStatistics -IncludeOldestAndNewestItems for the age of the oldest mail item, Export-MailboxDiagnosticLogs -ExtendedProperties for the ELC counters. Both sizes are shown with a meter of the archive share. Every blocker in Microsoft's MRM troubleshooting article becomes a finding with the article named and one line on why it matters. The oldest item is compared with the archive age: if nothing qualifies yet, the tool says so and gives the date the first item will.
3. **Prepare.** A panel lists exactly what will change: assign (or first create) the policy, enable the archive (or reconnect a disabled one; a disabled archive older than 30 days is removed first because Exchange refuses to reconnect it), lift the retention hold or the ELC block when you asked for that, enable auto-expanding archive when you asked. One approval for the set. Each change is re-read from Exchange afterwards and recorded as Verified or NotVerified.
4. **Run and watch.** Before the first pass the tool records the baseline. Each pass calls Start-ManagedFolderAssistant (with -FullCrawl on the first pass when requested), then waits the interval. During the wait it re-reads both mailbox sizes about once a minute and prints every batch that has landed in the archive as it is seen, with the meter filling; the glow line shows the countdown to the next size check and to the results. At the end of the interval it reads the ELC counters and the sizes. A pass is Verified when `ElcLastRunArchivedFromRootItemCount` is above zero, when the archive item count grew, or when `ElcLastSuccessTimestamp` advanced (the assistant completed a fresh run and found nothing eligible). Two consecutive passes that complete cleanly and move nothing end the run early. With a 0-minute interval the pass is recorded as Started and not measured. The MRM diagnostic log is saved; "resource unhealthy" in it is reported as throttling, which is expected on large mailboxes.
5. **Report.** A "What this means" paragraph in plain words, then report.html, result.json, CASE-NOTES.txt, run.log, and one MRM log per mailbox, written in a finally block so an interrupted run still has them.

## Passes and the interval

Exchange runs the Managed Folder Assistant on its own at least once every 7 days; that is fixed. A pass is this tool asking Exchange to process the mailbox now, waiting, and measuring. The interval is the wait, because the request returns immediately and the move happens in the background. One request is enough for a normal mailbox. A large mailbox is throttled and moves in slices, so several requests spread over hours finish sooner than one. Menu presets: once with a 10-minute wait; 3 passes 10 minutes apart; 6 passes 15 minutes apart; 12 passes 30 minutes apart; start and leave; custom (1 to 48 passes, 0 to 240 minutes).

## Creating a policy

`New-RetentionPolicyTag -Type All -RetentionEnabled $true -AgeLimitForRetention <days> -RetentionAction MoveToArchive` creates the default move-to-archive tag; `New-RetentionPolicy -RetentionPolicyTagLinks` wraps it in the policy. A policy may hold only one default tag per action, so the tool never copies default policy tags from elsewhere; it carries over the personal and folder tags of the Default MRM Policy unless -ArchiveTagOnly is given. A tag with identical settings that already exists is reused. Names are limited to 64 characters. The exact commands are shown before anything is created and require YES.

## What it will not do

- Delete anything. There is no cmdlet in this tool that removes items, folders or mailboxes.
- Lift a retention hold or clear ElcProcessingDisabled without the matching switch. Both are often deliberate.
- Edit or delete existing retention tags or policies. It can create a new tag and policy when you ask, and it assigns policies; it never changes one that already exists.
- Touch accounts. A disabled account is reported with the fix, not enabled.
- Run when the mailbox is blocked. If the archive is not active or a hold is on after the prepare stage, the assistant is not started, because the result would be nothing and the report would be misleading.

## Sizes

Exchange returns sizes as ByteQuantifiedSize objects in some sessions and as text such as "1.5 GB (1,610,612,736 bytes)" in REST sessions. Both are parsed to bytes. Ages such as AgeLimitForRetention arrive as TimeSpan or as text and are handled the same way.

## Compatibility

Windows PowerShell 5.1 and PowerShell 7 with ExchangeOnlineManagement 3.x. Strict mode is on throughout. The console layer degrades to plain text when there is no console, when output is redirected, or when NO_COLOR is set.
