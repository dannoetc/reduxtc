<#
.SYNOPSIS
  Find recent mail forwarding enablement across mailboxes in the last N hours (default 48).

.DESCRIPTION
  1) Searches the Unified Audit Log for Set-Mailbox and Inbox Rule changes that indicate forwarding.
  2) Focuses on events where forwarding *may have been enabled* in the last N hours.
  3) For each affected mailbox, fetches current forwarding properties and inbox rules.
  4) Flags external forwarding targets and suspicious patterns.
  5) Writes a CSV report.

.PARAMETER HoursBack
  Lookback window in hours. Default: 48.

.PARAMETER UsersCsv
  Optional CSV with header 'UserPrincipalName' to scope search.

.PARAMETER OutputPath
  Output CSV path. Default: .\RecentForwardingEvents.csv
#>

[CmdletBinding()]
param(
    [int]$HoursBack = 48,
    [string]$UsersCsv,
    [string]$OutputPath = ".\RecentForwardingEvents.csv"
)

# ------------------------ Helpers ------------------------
function Write-Log {
    param(
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO',
        [string]$Message
    )
    $ts = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    Write-Host "[$ts] [$Level] $Message"
}

function Join-Values {
    param([Parameter(ValueFromPipeline=$true)]$Values)
    process {
        if ($null -eq $Values) { return $null }
        if ($Values -is [System.Collections.IEnumerable] -and -not ($Values -is [string])) {
            ($Values | ForEach-Object { $_ }) -join '; '
        } else { [string]$Values }
    }
}

# Determine tenant accepted domains to detect external
$acceptedDomains = @()

# ------------------------ Connect ------------------------
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Log -Message "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop

    Write-Log -Message "Connecting to Compliance/Audit..."
    Connect-IPPSSession -ErrorAction Stop | Out-Null
}
catch {
    Write-Log -Level ERROR -Message "Connection failed: $($_.Exception.Message)"
    throw
}

try {
    $acceptedDomains = (Get-AcceptedDomain -ErrorAction Stop).DomainName | ForEach-Object { $_.ToLower() }
    Write-Log -Message "Accepted domains: $((($acceptedDomains) -join ', '))"
}
catch {
    Write-Log -Level WARN -Message "Could not read accepted domains. External detection may be limited. $($_.Exception.Message)"
    $acceptedDomains = @()
}

function Test-ExternalAddress {
    param([string]$Address)
    if ([string]::IsNullOrWhiteSpace($Address)) { return $false }
    if ($acceptedDomains.Count -eq 0) { return $null } # unknown
    $domain = ($Address -split '@')[-1].ToLower()
    return -not ($acceptedDomains -contains $domain)
}

# ------------------------ Scope (optional) ------------------------
$userScope = $null
if ($UsersCsv -and (Test-Path $UsersCsv)) {
    try {
        $userScope = Import-Csv -Path $UsersCsv | Where-Object { $_.UserPrincipalName } | Select-Object -ExpandProperty UserPrincipalName -Unique
        Write-Log -Message "Scoping to $(($userScope | Measure-Object).Count) user(s) from CSV."
    }
    catch {
        Write-Log -Level WARN -Message "Failed to parse UsersCsv. Continuing without scope. $($_.Exception.Message)"
        $userScope = $null
    }
}

# ------------------------ Audit query ------------------------
$start = (Get-Date).AddHours(-1 * $HoursBack)
$end   = Get-Date
Write-Log -Message "Searching Unified Audit Log: $start to $end"

# We care about Set-Mailbox (forwarding props) and Inbox Rule changes that add forwarding
$opsMailbox = @('Set-Mailbox')
$opsRules   = @('New-InboxRule','Set-InboxRule','UpdateInboxRules')

$auditMailboxParams = @{
    StartDate = $start
    EndDate   = $end
    Operations = $opsMailbox
    ResultSize = 5000
}
$auditRulesParams = @{
    StartDate = $start
    EndDate   = $end
    Operations = $opsRules
    ResultSize = 5000
}

$auditMailbox = @()
$auditRules   = @()

if ($userScope) {
    foreach ($u in $userScope) {
        $a = Search-UnifiedAuditLog @auditMailboxParams -UserIds $u -ErrorAction SilentlyContinue
        if ($a) { $auditMailbox += $a }
        $r = Search-UnifiedAuditLog @auditRulesParams -UserIds $u -ErrorAction SilentlyContinue
        if ($r) { $auditRules += $r }
    }
} else {
    $auditMailbox = Search-UnifiedAuditLog @auditMailboxParams -ErrorAction SilentlyContinue
    $auditRules   = Search-UnifiedAuditLog @auditRulesParams -ErrorAction SilentlyContinue
}

# Filter mailbox events to those that touched forwarding-related properties
$forwardingMailboxEvents = @()
foreach ($rec in $auditMailbox) {
    try {
        $data = $rec.AuditData | ConvertFrom-Json
    } catch { $data = $null }
    $hit = $false
    $props = @{}
    if ($data -and $data.Parameters) {
        foreach ($p in $data.Parameters) { $props[$p.Name] = $p.Value }
        foreach ($name in @('ForwardingSmtpAddress','ForwardingSMTPAddress','ForwardingAddress','DeliverToMailboxAndForward')) {
            if ($props.ContainsKey($name)) { $hit = $true; break }
        }
    }
    if ($hit) { $forwardingMailboxEvents += $rec }
}

# Filter inbox rule events to ones that include forward/redirect actions
$forwardingRuleEvents = @()
foreach ($rec in $auditRules) {
    try {
        $data = $rec.AuditData | ConvertFrom-Json
    } catch { $data = $null }

    $hit = $false
    if ($data -and $data.Parameters) {
        $paramMap = @{}
        foreach ($p in $data.Parameters) { $paramMap[$p.Name] = $p.Value }

        # Look for common action flags in parameters or details
        $text = (($paramMap.Keys + $paramMap.Values) -join ' ').ToLower()
        if ($text -match 'forward' -or $text -match 'redirect') { $hit = $true }
    }

    if ($hit) { $forwardingRuleEvents += $rec }
}

Write-Log -Message "Mailbox forwarding events: $($forwardingMailboxEvents.Count)"
Write-Log -Message "Rule-based forwarding events: $($forwardingRuleEvents.Count)"

# ------------------------ Build unique user set ------------------------
$affectedUpns = @()
if ($forwardingMailboxEvents) { $affectedUpns += $forwardingMailboxEvents.UserIds }
if ($forwardingRuleEvents)   { $affectedUpns += $forwardingRuleEvents.UserIds }
$affectedUpns = $affectedUpns | Sort-Object -Unique

if ($affectedUpns.Count -eq 0) {
    Write-Log -Level WARN -Message "No forwarding enablement activity found in the last $HoursBack hour(s)."
}

# ------------------------ Pull live state per mailbox ------------------------
$results = @()

foreach ($upn in $affectedUpns) {
    $liveMailbox = $null
    $inboxRules  = @()
    $fwSmtp      = $null
    $fwAddr      = $null
    $deliverCopy = $null

    try {
        $liveMailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
        $fwSmtp      = [string]$liveMailbox.ForwardingSmtpAddress
        $fwAddr      = $null
        if ($liveMailbox.ForwardingAddress) {
            # May be a recipient object
            $fwAddr = [string]$liveMailbox.ForwardingAddress
            if (-not $fwAddr) {
                $fwAddr = [string]$liveMailbox.ForwardingAddress.PrimarySmtpAddress
            }
        }
        $deliverCopy = [bool]$liveMailbox.DeliverToMailboxAndForward
    }
    catch {
        Write-Log -Level WARN -Message "Get-Mailbox failed for $upn $($_.Exception.Message)"
    }

    try {
        $inboxRules = Get-InboxRule -Mailbox $upn -ErrorAction Stop
    }
    catch {
        Write-Log -Level WARN -Message "Get-InboxRule failed for $upn $($_.Exception.Message)"
    }

    # Current rule actions (forward/redirect)
    $ruleTargets = @()
    $ruleNames   = @()
    foreach ($r in $inboxRules) {
        $targets = @()

        if ($r.ForwardTo) {
            foreach ($t in $r.ForwardTo) {
                $addr = $null
                foreach ($cand in @('PrimarySmtpAddress','ExternalEmailAddress','WindowsEmailAddress','Address','SmtpAddress')) {
                    if ($t.PSObject.Properties[$cand]) { $addr = [string]$t.$cand; if ($addr) { break } }
                }
                if (-not $addr) { $addr = [string]$t }
                if ($addr) { $targets += $addr }
            }
        }
        if ($r.RedirectTo) {
            foreach ($t in $r.RedirectTo) {
                $addr = $null
                foreach ($cand in @('PrimarySmtpAddress','ExternalEmailAddress','WindowsEmailAddress','Address','SmtpAddress')) {
                    if ($t.PSObject.Properties[$cand]) { $addr = [string]$t.$cand; if ($addr) { break } }
                }
                if (-not $addr) { $addr = [string]$t }
                if ($addr) { $targets += $addr }
            }
        }

        if ($targets.Count -gt 0) {
            $ruleTargets += $targets
            $ruleNames   += $r.Name
        }
    }

    # Gather audit snippets for this user
    $myMailboxEvents = $forwardingMailboxEvents | Where-Object { $_.UserIds -eq $upn }
    $myRuleEvents    = $forwardingRuleEvents   | Where-Object { $_.UserIds -eq $upn }

    foreach ($rec in ($myMailboxEvents + $myRuleEvents)) {
        $data = $null
        try { $data = $rec.AuditData | ConvertFrom-Json } catch {}

        # Build a compact param summary
        $paramSummary = $null
        if ($data -and $data.Parameters) {
            $pairs = @()
            foreach ($p in $data.Parameters) {
                $pairs += ("{0}={1}" -f $p.Name, $p.Value)
            }
            $paramSummary = ($pairs -join '; ')
        }

        # Determine suspicion
        $susp = @()

        # Mailbox-level forwarding now
        $currentTargets = @()
        if ($fwSmtp) { $currentTargets += $fwSmtp }
        if ($fwAddr) { $currentTargets += $fwAddr }

        $externalFlags = @()
        foreach ($t in $currentTargets + $ruleTargets) {
            $isExt = Test-ExternalAddress -Address $t
            if ($isExt -eq $true) { $externalFlags += $t }
        }

        if ($externalFlags.Count -gt 0) { $susp += "External forwarding target(s): $((($externalFlags) -join '; '))" }
        if ($deliverCopy) { $susp += "DeliverToMailboxAndForward=True" }
        if ($ruleTargets.Count -gt 0) { $susp += "Inbox rule forwarding active (rules: $((($ruleNames | Sort-Object -Unique) -join ', ')))" }

        $results += [pscustomobject]@{
            TimeUTC               = [datetime]::SpecifyKind($rec.CreationDate, [System.DateTimeKind]::Utc)
            Operation             = $rec.Operations
            UserPrincipalName     = $upn
            AuditParamSummary     = $paramSummary
            Current_ForwardingSMTPAddress = $fwSmtp
            Current_ForwardingAddress     = $fwAddr
            Current_DeliverToMailboxAndForward = $deliverCopy
            Current_RuleForwardTargets    = ($ruleTargets | Sort-Object -Unique | Join-Values)
            SuspiciousIndicators   = ($susp | Join-Values)
            ClientIP               = if ($data) { $data.ClientIP } else { $null }
            UserAgent              = if ($data) { $data.UserAgent } else { $null }
        }
    }
}

# ------------------------ Output ------------------------
try {
    $results | Sort-Object TimeUTC -Descending | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "Report written: $OutputPath"
}
catch {
    Write-Log -Level ERROR -Message "Failed writing CSV: $($_.Exception.Message)"
    throw
}

# ------------------------ Console summary ------------------------
if ($results.Count -gt 0) {
    Write-Host ""
    Write-Log -Message "Forwarding-related events (last $HoursBack hours):"
    $results |
        Where-Object { $_.Current_ForwardingSMTPAddress -or $_.Current_ForwardingAddress -or $_.Current_RuleForwardTargets -or $_.AuditParamSummary } |
        Select-Object TimeUTC, UserPrincipalName, Operation, SuspiciousIndicators |
        Format-Table -AutoSize
} else {
    Write-Log -Message "No results."
}
