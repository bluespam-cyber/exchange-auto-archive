# Contributing

Short rules, because this tool changes settings on other people's mailboxes.

- Every check or change must point to a Microsoft article, listed in docs/sources.md and named in the finding's Source.
- No deletions, ever. If a future feature needs to remove something, it does not belong in this tool.
- Anything that can be deliberate in a tenant (holds, ElcProcessingDisabled, policy contents, accounts) is reported, and changed only behind an explicit switch plus approval.
- Every change is re-read from Exchange afterwards and recorded as Verified or NotVerified. "The cmdlet returned" is not verification.
- Plain sentences, no emoji, no em dashes. Strict mode stays on; guard optional properties.
