# Sources

Every check, blocker and repair in Invoke-MailboxArchive.ps1 traces to one of these Microsoft articles.

## Archive mailboxes
- [Enable archive mailboxes for Microsoft 365](https://learn.microsoft.com/en-us/purview/enable-archive-mailboxes): Enable-Mailbox -Archive, the Mail Recipients role, the 30-day reconnect window for a disabled archive and the Set-Mailbox -RemoveDisabledArchive step, the automatic archive provisioning setting.
- [Enable auto-expanding archiving](https://learn.microsoft.com/en-us/purview/enable-autoexpanding-archiving): Enable-Mailbox -AutoExpandingArchive and its constraints.

## Retention policies and the Managed Folder Assistant
- [Customize an archive and deletion policy (MRM) for mailboxes](https://learn.microsoft.com/en-us/purview/set-up-an-archive-and-deletion-policy-for-mailboxes): a mailbox has one MRM policy, the assistant processes mailboxes at least once every seven days, Start-ManagedFolderAssistant, Set-Mailbox -ElcProcessingDisabled, items are not moved for a disabled account, Set-MailboxPlan for new mailboxes.
- [Start-ManagedFolderAssistant](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/start-managedfolderassistant): parameters including -FullCrawl, -HoldCleanup, -InactiveMailbox; the advice to use the mailbox GUID as the identity when an error occurs.
- [Retention tags and retention policies in Exchange Online](https://learn.microsoft.com/en-us/exchange/security-and-compliance/messaging-records-management/retention-tags-and-policies): default policy tags, retention policy tags, personal tags and how they combine; one default tag with the Move to Archive action per policy; the Default MRM Policy contents.
- [New-RetentionPolicyTag](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-retentionpolicytag): -Type All for a default policy tag, -RetentionAction MoveToArchive, -AgeLimitForRetention in days, -RetentionEnabled.
- [New-RetentionPolicy](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-retentionpolicy): -RetentionPolicyTagLinks.
- [Get-MailboxFolderStatistics](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-mailboxfolderstatistics): -IncludeOldestAndNewestItems, used to compare the oldest mail item with the archive age.
- [Feature permissions in Exchange Online](https://learn.microsoft.com/en-us/exchange/permissions-exo/feature-permissions): Messaging records management needs Compliance Management, Organization Management or Records Management; retention policy assignment needs Organization Management, Recipient Management or Records Management.

## Troubleshooting
- [Resolve email archive and deletion issues when using retention policies](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365/purview/retention/troubleshoot-mrm-email-archive-deletion): the 10 MB minimum, RetentionHoldEnabled, ElcProcessingDisabled, disabled tags, personal-only policies, the ELC counters through Export-MailboxDiagnosticLogs -ExtendedProperties, the MRM component log and the "resource unhealthy" throttling message, the FullCrawl step for a stale hidden MRM configuration.
- [Difference between ElcProcessingDisabled and RetentionHoldEnabled](https://learn.microsoft.com/en-us/troubleshoot/exchange/retention-policies/difference-between-elcprocessingdisabled-and-retentionholdenabled).

## Module
- [Connect to Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell): Connect-ExchangeOnline, Get-ConnectionInformation, module 3.x.
