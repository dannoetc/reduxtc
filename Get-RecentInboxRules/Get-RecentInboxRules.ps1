<#
.SYNOPSIS
  Report recent Inbox Rule activity (created/modified/removed) and flag risky rules.

.DESCRIPTION
  1) Queries the Unified Audit Log for Inbox rule operations in the last N days.
  2) For each affected mailbox, gets current Inbox rules.
  3) Flags suspicious actions (external forwarding/redirect, delete+stop, etc.).
  4) Writes a CSV report.

.PARAMETER DaysBack
  Number of days back to search (default 7).

.PARAMETER UsersCsv
  Optional path to a CSV with a header "UserPrincipalName" to scope the search to specific users.

.PARAMETER OutputPath
  Output CSV file path (default: .\RecentInboxRules.csv)

.EXAMPLE
  .\Get-RecentInboxRules.ps1 -DaysBack 14 -OutputPath C:\Temp\InboxRules.csv

.EXAMPLE
  .\Get-RecentInboxRules.ps1 -UsersCsv .\users.csv
#>

[CmdletBinding()]
param(
    [int]$DaysBack = 7,
    [string]$UsersCsv,
    [string]$OutputPath = ".\RecentInboxRules.csv"
)

#---------------------------#
# Helper: Write-Log
#---------------------------#
function Write-Log {
    param(
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO',
        [string]$Message
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    Write-Host "[$timestamp] [$Level] $Message"
}

#---------------------------#
# Helper: Safe property join
#---------------------------#
function Join-Values {
    param(
        [Parameter(ValueFromPipeline=$true)]
        $Values
    )
    process {
        if ($null -eq $Values) { return $null }
        if ($Values -is [System.Collections.IEnumerable] -and -not ($Values -is [string])) {
            ($Values | ForEach-Object { $_ }) -join '; '
        } else {
            [string]$Values
        }
    }
}

#---------------------------#
# Connect to EXO + Compliance (Audit)
#---------------------------#
try {
    Write-Log -Message "Connecting to Exchange Online..."
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop

    Write-Log -Message "Connecting to Compliance/Audit (Search-UnifiedAuditLog)..."
    # In v3+, this uses the same module. Command will auto-wire an IPPSSession.
    Connect-IPPSSession -ErrorAction Stop | Out-Null
}
catch {
    Write-Log -Level ERROR -Message "Failed to connect: $($_.Exception.Message)"
    throw
}

#---------------------------#
# Scope users (optional)
#---------------------------#
$userFilter = $null
if ($UsersCsv -and (Test-Path $UsersCsv)) {
    try {
        $userFilter = Import-Csv -Path $UsersCsv | Where-Object { $_.UserPrincipalName } | Select-Object -ExpandProperty UserPrincipalName -Unique
        Write-Log -Message "Scoping to $(($userFilter | Measure-Object).Count) user(s) from CSV."
    }
    catch {
        Write-Log -Level WARN -Message "Failed to read UsersCsv '$UsersCsv': $($_.Exception.Message). Continuing without user filter."
        $userFilter = $null
    }
}

#---------------------------#
# Accepted domains (for external test)
#---------------------------#
$acceptedDomains = @()
try {
    $acceptedDomains = (Get-AcceptedDomain -ErrorAction Stop).DomainName
    Write-Log -Message "Found accepted domains: $((($acceptedDomains) -join ', '))"
}
catch {
    Write-Log -Level WARN -Message "Couldn't fetch accepted domains; external detection may be limited. $($_.Exception.Message)"
}

function Is-ExternalAddress {
    param([string]$EmailAddress)
    if ([string]::IsNullOrWhiteSpace($EmailAddress)) { return $false }
    $domain = $EmailAddress.Split('@')[-1]
    if ([string]::IsNullOrWhiteSpace($domain)) { return $false }
    # Compare to accepted domains (case-insensitive)
    return -not ($acceptedDomains -contains $domain.ToLower())
}

#---------------------------#
# Audit search
#---------------------------#
$start = (Get-Date).AddDays(-1 * $DaysBack)
$end   = Get-Date
Write-Log -Message "Searching Unified Audit Log from $start to $end..."

# Operations of interest
$ops = @(
    "New-InboxRule",      # user-created
    "Set-InboxRule",      # user/administrator modified
    "Remove-InboxRule",   # deleted
    "UpdateInboxRules"    # Outlook/OWA batch updates
)

# Build base params
$auditParams = @{
    StartDate = $start
    EndDate   = $end
    Operations = $ops
    ResultSize = 5000
}

if ($userFilter) {
    # We'll loop per user when filter is provided to avoid missing matches due to server-side limits
    $auditRecords = @()
    foreach ($u in $userFilter) {
        Write-Log -Message "Querying audit for $u ..."
        $records = Search-UnifiedAuditLog @auditParams -UserIds $u -ErrorAction SilentlyContinue
        if ($records) { $auditRecords += $records }
    }
} else {
    $auditRecords = Search-UnifiedAuditLog @auditParams -ErrorAction SilentlyContinue
}

if (-not $auditRecords) {
    Write-Log -Level WARN -Message "No inbox rule activity found in last $DaysBack day(s)."
}

#---------------------------#
# Build index of mailboxes touched
#---------------------------#
$touchedUsers = @()
if ($auditRecords) {
    $touchedUsers = $auditRecords.UserIds | Sort-Object -Unique
    Write-Log -Message "Found $($touchedUsers.Count) mailbox(es) with recent rule activity."
}

#---------------------------#
# Pull live rules per mailbox
#---------------------------#
$liveRulesByUpn = @{}
foreach ($upn in $touchedUsers) {
    try {
        Write-Log -Message "Fetching current rules for $upn ..."
        $rules = Get-InboxRule -Mailbox $upn -ErrorAction Stop
        $liveRulesByUpn[$upn.ToLower()] = $rules
    }
    catch {
        Write-Log -Level WARN -Message "Could not get rules for $upn $($_.Exception.Message)"
    }
}

#---------------------------#
# Build results
#---------------------------#
$results = @()

foreach ($rec in $auditRecords) {
    # Parse AuditData JSON for context (rule name, parameters)
    $data = $null
    try {
        $data = $rec.AuditData | ConvertFrom-Json
    }
    catch {
        # skip malformed
    }

    $op          = $rec.Operations
    $user        = $rec.UserIds
    $when        = $rec.CreationDate
    $ruleName    = $null
    $clientIP    = $data.ClientIP
    $userAgent   = $data.UserAgent
    $paramsTable = @{}

    if ($data -and $data.Parameters) {
        foreach ($p in $data.Parameters) {
            $paramsTable[$p.Name] = $p.Value
        }
        if ($paramsTable.ContainsKey("Name")) { $ruleName = $paramsTable["Name"] }
        elseif ($paramsTable.ContainsKey("RuleName")) { $ruleName = $paramsTable["RuleName"] }
    }

    $liveRules = $liveRulesByUpn[$user.ToLower()]
    $matchingRule = $null
    if ($liveRules -and $ruleName) {
        $matchingRule = $liveRules | Where-Object { $_.Name -eq $ruleName } | Select-Object -First 1
    }

    # Extract live rule details (if present)
    $enabled        = $null
    $priority       = $null
    $conditions     = $null
    $actions        = $null
    $externalTargets = @()
    $suspicion      = @()

    if ($matchingRule) {
        $enabled  = $matchingRule.Enabled
        $priority = $matchingRule.Priority

        # Conditions summary
        $condParts = @()
        foreach ($prop in 'From','FromAddressContainsWords','SubjectContainsWords','SubjectOrBodyContainsWords','MyNameInToOrCcBox','SentTo') {
            $val = $matchingRule.$prop
            if ($null -ne $val -and $val -ne $false) {
                $condParts += "$prop $((Join-Values $val))"
            }
        }
        $conditions = $condParts -join ' | '

        # Actions summary + external check
        $actParts = @()
        foreach ($prop in 'DeleteMessage','MarkAsRead','StopProcessingRules','MoveToFolder','ForwardTo','RedirectTo','CopyToFolder','PermanentDelete') {
            $val = $matchingRule.$prop
            if ($null -ne $val -and $val -ne $false) {
                if ($prop -in @('ForwardTo','RedirectTo')) {
                    # These are Recipient objects; pull PrimarySmtpAddress or ToString()
                    $recips = @()
                    foreach ($r in $val) {
                        # Try common properties
                        $addr = $null
                        foreach ($candidate in @('PrimarySmtpAddress','ExternalEmailAddress','WindowsEmailAddress','Address','SmtpAddress')) {
                            if ($r.PSObject.Properties[$candidate]) {
                                $addr = [string]$r.$candidate
                                if ($addr) { break }
                            }
                        }
                        if (-not $addr) { $addr = [string]$r }
                        $recips += $addr
                        if (Is-ExternalAddress -EmailAddress $addr) {
                            $externalTargets += $addr
                        }
                    }
                    $actParts += "$prop $((Join-Values $recips))"
                } else {
                    $actParts += "$prop $((Join-Values $val))"
                }
            }
        }
        $actions = $actParts -join ' | '

        if ($externalTargets.Count -gt 0) {
            $suspicion += "External $($prop -replace 'To','') target(s): $((Join-Values $externalTargets))"
        }
        if ($matchingRule.DeleteMessage -and $matchingRule.StopProcessingRules) {
            $suspicion += "Delete + StopProcessingRules"
        }
        if ($matchingRule.MoveToFolder -and $matchingRule.StopProcessingRules) {
            $suspicion += "Move ($($matchingRule.MoveToFolder)) + StopProcessingRules"
        }
        if ($matchingRule.MarkAsRead -and $matchingRule.StopProcessingRules) {
            $suspicion += "MarkAsRead + StopProcessingRules"
        }
    } else {
        # Rule missing could mean removed post-event
        if ($op -eq 'Remove-InboxRule') {
            $suspicion += "Rule was removed"
        } elseif ($ruleName) {
            $suspicion += "Rule '$ruleName' not found now (possibly removed)"
        }
    }

    $results += [pscustomobject]@{
        TimeUTC              = [datetime]::SpecifyKind($when, [System.DateTimeKind]::Utc)
        Operation            = $op
        UserPrincipalName    = $user
        RuleName             = $ruleName
        Enabled              = $enabled
        Priority             = $priority
        Conditions           = $conditions
        Actions              = $actions
        ExternalTargets      = ($externalTargets | Join-Values)
        SuspiciousIndicators = ($suspicion | Join-Values)
        ClientIP             = $clientIP
        UserAgent            = $userAgent
    }
}

#---------------------------#
# Write CSV
#---------------------------#
try {
    $results | Sort-Object TimeUTC -Descending | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $OutputPath
    Write-Log -Message "Report written to $OutputPath"
}
catch {
    Write-Log -Level ERROR -Message "Failed to write CSV: $($_.Exception.Message)"
    throw
}

#---------------------------#
# Extra: summarize to screen
#---------------------------#
if ($results.Count -gt 0) {
    Write-Host ""
    Write-Log -Message "Top suspicious hits:"
    $results |
        Where-Object { $_.SuspiciousIndicators } |
        Select-Object TimeUTC, UserPrincipalName, RuleName, SuspiciousIndicators |
        Format-Table -AutoSize
} else {
    Write-Log -Message "No results to report."
}
