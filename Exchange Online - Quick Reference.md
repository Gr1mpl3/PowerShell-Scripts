# Exchange Online PowerShell — Quick Reference

## Connect to EXO PowerShell

```powershell
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.onmicrosoft.com
```

---

## Groups

### Confirm group members / owners

```powershell
Get-UnifiedGroupLinks -Identity "GroupName" | Format-List *
```

> Add `-LinkType Members` or `-LinkType Owners` to filter to just one type instead of pulling all link types.

### Confirm group info

```powershell
Get-UnifiedGroup -Identity "GroupName" | Format-List *
```

---

## Mailboxes

### Check mailbox creation date

```powershell
Get-Mailbox -Identity "user@domain.com" | Select-Object WhenCreated
```

---

## Inbox Rules

### Show all rules (including hidden)

```powershell
Get-InboxRule -Mailbox "user@domain.com" -IncludeHidden | Select-Object *
```

### Disable a specific rule

```powershell
Disable-InboxRule -Mailbox "user@domain.com" -Identity "RuleName"
```

### Remove a specific rule

```powershell
Remove-InboxRule -Mailbox "user@domain.com" -Identity "RuleName"
```

### Remove all rules

```powershell
Get-InboxRule -Mailbox "user@domain.com" | Remove-InboxRule
```

> No confirmation prompt by default in some sessions — double-check the mailbox before running on a live account.

---

## Mail Forwarding & Delegate Access

> Useful for BEC/compromised-account triage — attackers commonly add a forwarding rule or delegate to silently exfiltrate mail.

### Check mailbox-level forwarding (not a rule — the ForwardingSMTPAddress property)

```powershell
Get-Mailbox -Identity "user@domain.com" | Select-Object ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward
```

### Remove mailbox-level forwarding

```powershell
Set-Mailbox -Identity "user@domain.com" -ForwardingSmtpAddress $null -ForwardingAddress $null
```

### Check mailbox permissions (delegates / full access grants)

```powershell
Get-MailboxPermission -Identity "user@domain.com" | Where-Object { $_.User -notlike "NT AUTHORITY\SELF" }
```

### Check "Send As" / "Send on Behalf" delegates

```powershell
Get-RecipientPermission -Identity "user@domain.com"
Get-Mailbox -Identity "user@domain.com" | Select-Object GrantSendOnBehalfTo
```

---

## Mail Flow / Message Trace

### Trace recent messages (last 10 days max via this cmdlet)

```powershell
Get-MessageTrace -SenderAddress "user@domain.com" -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date)
```

### Trace by recipient

```powershell
Get-MessageTrace -RecipientAddress "user@domain.com" -StartDate (Get-Date).AddDays(-3) -EndDate (Get-Date)
```

> For anything older than 10 days, use the Defender/Exchange admin center's extended message trace, or `Get-MessageTraceDetail` for a specific message's routing hops.

---

## Mailbox Audit Logging

> Worth checking first on any "did someone access this mailbox" ticket — if audit logging was off, you may have no trail.

### Confirm audit logging is enabled

```powershell
Get-Mailbox -Identity "user@domain.com" | Select-Object AuditEnabled, AuditLogAgeLimit
```

### Search the audit log for a mailbox

```powershell
Search-MailboxAuditLog -Identity "user@domain.com" -LogonTypes Owner,Delegate,Admin -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date)
```

---

## Account / Sign-In Status

### Check mailbox sync status

```powershell
Get-EXOMailbox -Identity "user@domain.com" | Select-Object DisplayName, IsDirSynced
```

> Note: Account enable/disable and actual sign-in status live in Entra ID, not EXO — EXO only manages the mailbox object. Use `Get-MgUser -UserId "user@domain.com" | Select-Object AccountEnabled` (Microsoft Graph) instead.

### Disconnect a user's active sessions (force re-auth, e.g. after a password reset on a suspected compromise)

```powershell
Revoke-MgUserSignInSession -UserId "user@domain.com"
```

> This is a Microsoft Graph cmdlet, not EXO — needs `Connect-MgGraph -Scopes "User.RevokeSessions.All"` (or equivalent) in the same session.

---

## Distribution Groups (legacy, non-Microsoft 365 Groups)

### Confirm DL members

```powershell
Get-DistributionGroupMember -Identity "GroupName"
```

### Add/remove a member

```powershell
Add-DistributionGroupMember -Identity "GroupName" -Member "user@domain.com"
Remove-DistributionGroupMember -Identity "GroupName" -Member "user@domain.com"
```