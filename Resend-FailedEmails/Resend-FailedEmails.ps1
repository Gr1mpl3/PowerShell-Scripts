<#
.SYNOPSIS
    Re-forwards previously delivered emails from a Microsoft 365 mailbox to a
    target address (such as a Help Scout inbound address) based on a CSV list.

.DESCRIPTION
    Reads a CSV of emails that need to be resent (for example, messages that
    failed to reach a help desk during a mail migration or routing change).
    For each row, the script searches the connected mailbox for a message that
    matches the subject within a +/- 6 hour window around the original
    timestamp. Any match is forwarded to the configured destination address.

    Work is done in batches with a pause between them to stay under Microsoft
    Graph throttling limits.

.NOTES
    Author : Gr1mpl3
    Version: 1.0

    See README.md for full setup, CSV format, and usage instructions.
#>

# ============================================================================
#  PREREQUISITES
# ----------------------------------------------------------------------------
#  1. Install the Microsoft Graph PowerShell SDK (one time):
#
#         Install-Module Microsoft.Graph -Scope CurrentUser
#
#  2. Sign in to Graph BEFORE running this script, using an account that owns
#     (or has delegated access to) the mailbox you want to process, with the
#     scopes needed to read and send mail:
#
#         Connect-MgGraph -Scopes "Mail.Read","Mail.Send"
#
#     This script operates on "/me" — the signed-in account's own mailbox.
#     Whatever account you run Connect-MgGraph with is the mailbox that gets
#     searched and the mailbox that sends the forwards.
# ============================================================================

Import-Module Microsoft.Graph.Authentication

# Make sure a Graph session already exists. We do NOT connect here on purpose,
# so the operator stays in control of which account/scopes are used.
if (-not (Get-MgContext)) {
    Write-Host "Not connected to Microsoft Graph. Run Connect-MgGraph first, then rerun." -ForegroundColor Red
    exit
}

# ============================================================================
#  CONFIGURATION  —  EDIT THESE VALUES BEFORE RUNNING
# ============================================================================

# (Informational) The mailbox you connected as. The Graph calls below use
# "/me", so this value is for your own reference and is not used directly.
#
# TIP: To target a DIFFERENT mailbox without signing in as it, replace the
#      "/me/" segments further down with "/users/$Mailbox/" and connect with
#      an account or app registration that has the appropriate delegated or
#      application permissions to that mailbox.
$Mailbox     = "shared-mailbox@yourdomain.com"

# The destination address that every matched message is forwarded to.
# For Help Scout this is the inbound/forwarding address found under
# Mailbox > Settings > Emails. It looks like:
#     support.NNNNN.xxxxxxxxxxxxxxxx@helpscout.net
# This can be any valid email address if you are not using Help Scout.
$ForwardTo   = "support.00000.0000000000000000@helpscout.net"

# Full path to the CSV of emails to resend. See README.md for the required
# columns (message_subject, sender_address, date_time_utc).
$CsvPath     = "C:\path\to\Failed_Emails.csv"

# How many emails to process per batch before pausing.
$BatchSize    = 10

# How long (in seconds) to pause between batches, to avoid Graph throttling.
# 600 = 10 minutes. Lower this for small jobs; raise it if you hit 429 errors.
$PauseSeconds = 600

# SAFETY SWITCH:
#   $true  -> DRY RUN. The script only REPORTS what it would forward and sends
#             nothing. Leave this as $true for your first run.
#   $false -> LIVE. Emails ARE actually forwarded.
# Confirm the dry-run output looks correct, then set this to $false.
$DryRun = $true

# The note prepended to each forwarded message. Replace the signature.
$ForwardComment = @"
This is a resend of a previous email that did not go through due to a delivery issue.

Thank you,
[Your Company Name]
"@

# ============================================================================
#  END CONFIGURATION  —  no edits needed below this line for normal use
# ============================================================================

# --- Optional: limit to the first few rows for an initial production test. ---
# Uncomment the next line (and comment out the full-run line) to test on 3 rows:
# $Items = Import-Csv $CsvPath | Select-Object -First 3

# Full run: load every row from the CSV.
$Items = Import-Csv $CsvPath

$total = $Items.Count
$index = 0

Write-Host "Total emails to process: $total" -ForegroundColor Cyan

if ($DryRun) {
    Write-Host "DRY RUN MODE: No emails will be forwarded." -ForegroundColor Magenta
}
else {
    Write-Host "LIVE MODE: Emails WILL be forwarded." -ForegroundColor Red
}

# Process the CSV in batches so we can pause between groups and avoid throttling.
while ($index -lt $total) {

    $batch = $Items[$index..([math]::Min($index + $BatchSize - 1, $total - 1))]

    Write-Host "`nProcessing batch starting at index $index..." -ForegroundColor Yellow

    foreach ($item in $batch) {

        # Pull and clean the fields we match on from this CSV row.
        $subject = ($item.message_subject).Trim().Trim('"')
        $sender  = $item.sender_address
        $time    = [datetime]$item.date_time_utc

        # A blank subject can't be matched reliably, so skip it.
        if ([string]::IsNullOrWhiteSpace($subject)) {
            Write-Host "SKIPPING: Blank subject | From: $sender | Around: $time" -ForegroundColor DarkYellow
            continue
        }

        # Build a +/- 6 hour window around the original timestamp to search in.
        # Widen or narrow these offsets if your timestamps are less precise.
        $start = $time.AddHours(-6).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $end   = $time.AddHours(6).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

        try {
            # Query the connected mailbox for messages in the time window.
            # NOTE: change "/me/" to "/users/$Mailbox/" to target another mailbox.
            $uri = "https://graph.microsoft.com/v1.0/me/messages?`$filter=receivedDateTime ge $start and receivedDateTime le $end&`$select=id,subject,from,receivedDateTime,parentFolderId&`$top=100"

            $response = Invoke-MgGraphRequest -Method GET -Uri $uri
            $messages = $response.value

            Write-Host "  Messages found in time window: $($messages.Count)" -ForegroundColor DarkCyan

            # From the candidates, keep the newest message whose subject is an
            # exact (trimmed) match for the CSV subject.
            $match = $messages | Where-Object {
                $_.subject.Trim() -eq $subject
            } | Sort-Object receivedDateTime -Descending | Select-Object -First 1

            if ($match) {

                if ($DryRun) {
                    # Dry run: report the match but do not send anything.
                    Write-Host "WOULD FORWARD:" -ForegroundColor Green
                    Write-Host "  Subject: $($match.subject)"
                    Write-Host "  From:    $($match.from.emailAddress.address)"
                    Write-Host "  Date:    $($match.receivedDateTime)"
                    Write-Host "  To:      $ForwardTo"
                    Write-Host "  Message ID: $($match.id)"
                }
                else {
                    # Live: forward the matched message via the Graph forward action.
                    $forwardUri = "https://graph.microsoft.com/v1.0/me/messages/$($match.id)/forward"

                    $body = @{
                        comment      = $ForwardComment
                        toRecipients = @(
                            @{
                                emailAddress = @{
                                    address = $ForwardTo
                                }
                            }
                        )
                    } | ConvertTo-Json -Depth 5

                    Invoke-MgGraphRequest `
                        -Method POST `
                        -Uri $forwardUri `
                        -Body $body `
                        -ContentType "application/json"

                    Write-Host "FORWARDED:" -ForegroundColor Green
                    Write-Host "  Subject: $($match.subject)"
                    Write-Host "  From:    $($match.from.emailAddress.address)"
                    Write-Host "  Date:    $($match.receivedDateTime)"
                    Write-Host "  To:      $ForwardTo"
                }
            }
            else {
                Write-Host "NOT FOUND: $subject | From: $sender | Around: $time" -ForegroundColor Red
            }
        }
        catch {
            # Log and continue so one bad row doesn't stop the whole run.
            Write-Host "ERROR checking/sending: $subject" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }

    $index += $BatchSize

    # Pause before the next batch (skip the wait after the final batch).
    if ($index -lt $total) {
        Write-Host "Pausing $PauseSeconds seconds before next batch..." -ForegroundColor Cyan
        Start-Sleep -Seconds $PauseSeconds
    }
}
