<#
.SYNOPSIS
Report users who have NOT signed into OneDrive/SharePoint within a lookback window
using ONLY the Entra ID sign-in logs (fallback method).

.DESCRIPTION
- Pulls all users (enabled by default; include disabled with -IncludeDisabled).
- Queries sign-in logs for OneDrive/SharePoint apps within LookbackDays.
- Users with NO matching sign-ins are listed as "No OneDrive sign-ins (last N days)".

.REQUIREMENTS
- PowerShell 5.1+
- Modules: Microsoft.Graph.Reports, Microsoft.Graph.Users
- Graph scopes: AuditLog.Read.All, User.Read.All

.PARAMETERS
- LookbackDays    : Days of history to scan (default 180; note tenant retention limits)
- OutputPath      : Path for the HTML report (default in current folder)
- IncludeDisabled : Include disabled/blocked accounts in evaluation
- ExtraAppNames   : Additional appDisplayName values to consider as OneDrive/SharePoint sign-ins

.NOTES
If your sign-in log retention is shorter than LookbackDays, results reflect only retained data.
#>

[CmdletBinding()]
param(
  [int]$LookbackDays = 180,
  [string]$OutputPath = "$(Join-Path $PWD 'OneDrive-NeverSignedIn-Report.html')",
  [switch]$IncludeDisabled,
  [string[]]$ExtraAppNames
)

# --- Minimal modules ---
try { Import-Module Microsoft.Graph.Reports -ErrorAction Stop } catch { Write-Error "Missing Microsoft.Graph.Reports: $($_.Exception.Message)"; exit 1 }
try { Import-Module Microsoft.Graph.Users   -ErrorAction Stop } catch { Write-Error "Missing Microsoft.Graph.Users: $($_.Exception.Message)"; exit 1 }

# --- Connect with minimal scopes ---
if (-not (Get-MgContext)) {
  Connect-MgGraph -Scopes @('AuditLog.Read.All','User.Read.All') -NoWelcome
}

# --- Helper: HTML builder ---
function New-ReportHtml {
  param(
    [Parameter(Mandatory=$true)]$NeverRows,
    [Parameter(Mandatory=$true)]$SeenRows,
    [int]$LookbackDays,
    [datetime]$Generated = (Get-Date),
    [string]$Title = "Users With No OneDrive/SharePoint Sign-ins (last $LookbackDays days)"
  )
  $style = @"
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:24px}
h1{font-size:22px;margin:0 0 6px 0}
h2{font-size:18px;margin:20px 0 8px 0}
.small{color:#555}
table{border-collapse:collapse;width:100%;margin-top:8px}
th,td{border:1px solid #ddd;padding:8px} th{background:#f6f6f6;text-align:left}
tr:nth-child(even){background:#fafafa}
.badge{display:inline-block;padding:2px 8px;border-radius:12px;background:#eee;margin-left:6px}
.warn{background:#ffe8a6}
.ok{background:#d4f5d0}
</style>
"@
  $neverCount = $NeverRows.Count
  $seenCount  = $SeenRows.Count
  $summary = "<div class='small'><strong>Generated:</strong> $Generated<br/><strong>Window:</strong> last $LookbackDays days<br/><span class='badge warn'>No sign-in: $neverCount</span><span class='badge ok'>Seen: $seenCount</span></div>"
  $tNever = $NeverRows | Select-Object DisplayName,UserPrincipalName,AccountEnabled |
    ConvertTo-Html -Fragment -PreContent "<h2>No OneDrive/SharePoint sign-ins ($neverCount)</h2>"
  $tSeen  = $SeenRows | Select-Object DisplayName,UserPrincipalName,LastSignIn,App |
    ConvertTo-Html -Fragment -PreContent "<h2>Seen in OneDrive/SharePoint sign-ins ($seenCount)</h2>"

  ConvertTo-Html -Head $style -PreContent "<h1>$Title</h1>$summary" -PostContent ($tNever + $tSeen) -Title $Title
}

# --- Build user set (enabled by default) ---
$allUsers = Get-MgUser -All -Property "id,displayName,userPrincipalName,accountEnabled"
if (-not $IncludeDisabled) {
  $allUsers = $allUsers | Where-Object { $_.AccountEnabled -eq $true }
}
# map of UPN -> user object
$userMap = @{}
foreach ($u in $allUsers) {
  if ($u.UserPrincipalName) { $userMap[$u.UserPrincipalName.ToLower()] = $u }
}

# --- Query sign-in logs for OneDrive/SharePoint apps ---
$sinceIso = (Get-Date).AddDays(-[math]::Abs($LookbackDays)).ToString('o')

# Common appDisplayName values for OneDrive/SharePoint
$apps = @(
  "Office 365 SharePoint Online",
  "Microsoft SharePoint Online",
  "SharePoint Online",
  "Microsoft OneDrive",
  "OneDrive Sync Engine",
  "OneDrive Web"
)
if ($ExtraAppNames) { $apps += $ExtraAppNames }

# Build OData filter
$appFilter = ($apps | ForEach-Object { "appDisplayName eq '$_'" }) -join " or "
$filter = "(createdDateTime ge $sinceIso) and ($appFilter)"

# Pull sign-ins (Graph SDK handles paging with -All)
$rows = @()
try {
  $rows = Get-MgAuditLogSignIn -Filter $filter -All
} catch {
  Write-Error "Failed to query sign-in logs: $($_.Exception.Message)"
  exit 1
}

# --- Reduce to latest sign-in per UPN for these apps ---
$latestByUpn = @{}
foreach ($s in $rows) {
  $upn = $s.UserPrincipalName
  if ([string]::IsNullOrWhiteSpace($upn)) { continue }
  $key = $upn.ToLower()
  if (-not $userMap.ContainsKey($key)) {
    # skip users outside our selected set (e.g., disabled when not included)
    continue
  }
  $dt = $null
  try { $dt = [datetime]$s.CreatedDateTime } catch {}
  if (-not $dt) { continue }

  if (-not $latestByUpn.ContainsKey($key) -or $dt -gt $latestByUpn[$key].When) {
    $latestByUpn[$key] = [pscustomobject]@{
      DisplayName       = $s.UserDisplayName
      UserPrincipalName = $upn
      LastSignIn        = $dt.ToString("yyyy-MM-dd HH:mm:ss")
      App               = $s.AppDisplayName
      When              = $dt
    }
  }
}

# --- Partition users: seen vs never ---
$seen = New-Object System.Collections.Generic.List[Object]
$never = New-Object System.Collections.Generic.List[Object]

foreach ($kv in $userMap.GetEnumerator()) {
  $upn = $kv.Key
  $u   = $kv.Value
  if ($latestByUpn.ContainsKey($upn)) {
    $s = $latestByUpn[$upn]
    $seen.Add([pscustomobject]@{
      DisplayName       = if ($s.DisplayName) { $s.DisplayName } else { $u.DisplayName }
      UserPrincipalName = $s.UserPrincipalName
      LastSignIn        = $s.LastSignIn
      App               = $s.App
    })
  } else {
    $never.Add([pscustomobject]@{
      DisplayName       = $u.DisplayName
      UserPrincipalName = $u.UserPrincipalName
      AccountEnabled    = $u.AccountEnabled
    })
  }
}

# Sort for readability
$seen  = $seen  | Sort-Object DisplayName, UserPrincipalName
$never = $never | Sort-Object DisplayName, UserPrincipalName

# --- Emit HTML ---
$html = New-ReportHtml -NeverRows $never -SeenRows $seen -LookbackDays $LookbackDays
$html | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host "Report written to: $OutputPath"
Write-Host ("No OneDrive/SharePoint sign-ins: {0} | Seen: {1}" -f $never.Count, $seen.Count)
