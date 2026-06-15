# Resend-FailedEmails

A PowerShell utility that re-forwards previously delivered Microsoft 365 emails to a target address (for example, a [Help Scout](https://www.helpscout.com/) inbound address) based on a CSV list.

It's useful when a batch of emails failed to reach a help desk or shared inbox — during a mail migration, a routing/connector change, or an outage — and you have a message-trace export of what went missing. The script finds those messages in the mailbox and forwards them so they land where they should have the first time.

> **Note:** Forwarding through Microsoft Graph sends the message *from the connected mailbox*. Recipients will see it as a new email (with a short "this is a resend" note) rather than the original sender. For ticketing systems like Help Scout that key off subject and timestamp, this is usually fine, but confirm it fits your workflow before a live run.

---

## How it works

For each row in the CSV, the script:

1. Builds a **±6 hour time window** around the email's original timestamp.
2. Queries the connected mailbox (`/me`) for messages received in that window.
3. Keeps the **newest message whose subject exactly matches** the CSV subject.
4. **Forwards** that message to the configured destination address.

Work is done in **batches** with a configurable **pause** between them to stay under Microsoft Graph throttling limits. A **dry-run mode** lets you preview every match before anything is sent.

---

## Prerequisites

- **PowerShell 5.1+** or **PowerShell 7+**
- **Microsoft Graph PowerShell SDK**

  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```

- A Microsoft 365 account that **owns or has delegated access** to the mailbox you want to process, with permission to **read and send mail**.

---

## Permissions / scopes

Connect to Graph before running the script:

```powershell
Connect-MgGraph -Scopes "Mail.Read","Mail.Send"
```

| Scope       | Why it's needed                         |
|-------------|-----------------------------------------|
| `Mail.Read` | Search the mailbox for matching emails  |
| `Mail.Send` | Forward the matched emails              |

The script runs against `/me`, so **whichever account you sign in as is the mailbox that gets searched and the mailbox that sends the forwards.** To target a different mailbox, see [Targeting a different mailbox](#targeting-a-different-mailbox) below.

---

## CSV format

The CSV needs these columns (additional columns are ignored):

| Column            | Description                                  | Example                          |
|-------------------|----------------------------------------------|----------------------------------|
| `message_subject` | Subject line of the email                    | `Order confirmation #1234`       |
| `sender_address`  | Original sender (used for logging only)      | `customer@example.com`           |
| `date_time_utc`   | Timestamp used to build the search window    | `2025-01-15T14:32:00Z`           |

These column names line up with a **Microsoft 365 message trace export** (Exchange admin center / Defender → message trace → export). The typical workflow is: run a message trace for the affected period, filter to the failed/missing messages, export to CSV, and point the script at it.

---

## Configuration

Open `Resend-FailedEmails.ps1` and edit the values in the **CONFIGURATION** block near the top:

| Variable        | What to set it to                                                                 |
|-----------------|-----------------------------------------------------------------------------------|
| `$Mailbox`      | The mailbox you're connecting as (informational reference).                       |
| `$ForwardTo`    | Destination address. For Help Scout: **Mailbox → Settings → Emails** (`support.NNNNN.xxxx@helpscout.net`). |
| `$CsvPath`      | Full path to your CSV file.                                                        |
| `$BatchSize`    | Emails processed per batch (default `10`).                                         |
| `$PauseSeconds` | Seconds to wait between batches (default `600` = 10 min).                          |
| `$DryRun`       | `$true` to preview only, `$false` to actually forward. **Starts as `$true`.**     |
| `$ForwardComment` | The note prepended to each forwarded email — replace the signature.             |

---

## Usage

1. **Install the SDK** (one time):

   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```

2. **Connect** as the mailbox you want to process:

   ```powershell
   Connect-MgGraph -Scopes "Mail.Read","Mail.Send"
   ```

3. **Edit the configuration** block in the script.

4. **Dry run first.** With `$DryRun = $true`, run the script and review the `WOULD FORWARD` / `NOT FOUND` output:

   ```powershell
   .\Resend-FailedEmails.ps1
   ```

   For an even smaller first pass, uncomment the `Select-Object -First 3` line in the script to test on just three rows.

5. **Go live.** Once the dry-run output looks correct, set `$DryRun = $false` and run it again.

### Output legend

| Message         | Meaning                                                        |
|-----------------|----------------------------------------------------------------|
| `WOULD FORWARD` | Dry run found a match (nothing sent)                           |
| `FORWARDED`     | Live run forwarded the message                                 |
| `NOT FOUND`     | No subject match in the time window                            |
| `SKIPPING`      | Row had a blank subject                                        |
| `ERROR`         | A read/send error occurred (the run continues to the next row) |

---

## Targeting a different mailbox

By default the script uses `/me/`, the connected account's own mailbox. To process a mailbox you have delegated or application access to without signing in as it:

1. Connect with an account or app registration that has access to that mailbox.
2. In the script, replace the two `/me/` segments in the Graph URIs with `/users/$Mailbox/`.
3. Set `$Mailbox` to that mailbox's address or ID.

Application-permission access additionally requires an app registration with `Mail.Read` / `Mail.Send` granted and admin-consented.

---

## Limitations & notes

- **Subject matching is exact** (after trimming). Messages whose subjects were altered (e.g., a prefix added) won't match.
- **Duplicate subjects** in the same window: only the newest match is forwarded.
- **Time window** is ±6 hours. Adjust the `AddHours(-6)` / `AddHours(6)` values if your timestamps are off or imprecise.
- **Throttling:** if you see HTTP `429` responses, increase `$PauseSeconds` and/or lower `$BatchSize`.
- Forwarded messages are sent **from the connected mailbox**, not the original sender.
- Test against a non-production mailbox if you can before running at scale.

---

## Disclaimer

Provided as-is, with no warranty. You are responsible for how it's used in your environment. Always do a dry run, and verify you have authorization to access and forward mail in the target tenant/mailbox.

