<#
.SYNOPSIS
Generates an HTML report of which users have signed into / used OneDrive recently.
Skips blocked or disabled accounts (accountEnabled -eq $false).

.REQUIREMENTS
- Modules: Microsoft.Graph.Reports, Microsoft.Graph.Users
- Scopes: Reports.Read.All, AuditLog.Read.All, User.Read.All
#>

[CmdletBinding()]
param(
  [string]$OutputPath = "$(Join-Path $PWD 'OneDrive-SignIn-Report.html')",
  [ValidateSet('D7','D30','D90','D180')]
  [string]$Period = 'D30',
  [switch]$UseFallbackOnly
)

# ------------------ Imports ------------------
try { Import-Module Microsoft.Graph.Reports -ErrorAction Stop } catch { Write-Error "Missing Microsoft.Graph.Reports: $($_.Exception.Message)"; exit 1 }
try { Import-Module Microsoft.Graph.Users   -ErrorAction Stop } catch { Write-Error "Missing Microsoft.Graph.Users: $($_.Exception.Message)"; exit 1 }

# ------------------ Connect ------------------
$scopes = @('Reports.Read.All','AuditLog.Read.All','User.Read.All')
if (-not (Get-MgContext)) {
  Connect-MgGraph -Scopes $scopes -NoWelcome
}

# ------------------ Helpers ------------------
function New-ReportHtml {
  param($ActiveRows,$NoActivityRows,$FallbackRows,[string]$Title,[datetime]$Generated=(Get-Date))
  $style = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; }
h1 { font-size: 22px; margin-bottom: 4px; }
h2 { font-size: 18px; margin-top: 24px; }
table { border-collapse: collapse; width: 100%; margin-top: 8px; }
th,td { border:1px solid #ddd; padding:8px; }
th { background:#f6f6f6; }
tr:nth-child(even){background:#fafafa;}
</style>
"@
  $activeCount   = if ($ActiveRows)   { $ActiveRows.Count }   else { 0 }
  $noActivityCnt = if ($NoActivityRows){ $NoActivityRows.Count } else { 0 }
  $fallbackCount = if ($FallbackRows) { $FallbackRows.Count } else { 0 }

  $summary = "<p><strong>Generated:</strong> $Generated<br/><strong>Period:</strong> $Period<br/>" +
             "<strong>Active:</strong> $activeCount | " +
             "No Activity: $noActivityCnt" +
             ($(if ($FallbackRows) { " | Fallback: $fallbackCount" } else { "" })) + "</p>"

  $a = $ActiveRows | Select-Object DisplayName,UserPrincipalName,LastActivityDate,OneDriveUrl | 
       ConvertTo-Html -Fragment -PreContent "<h2>Active Users ($activeCount)</h2>"
  $n = $NoActivityRows | Select-Object DisplayName,UserPrincipalName,OneDriveUrl |
       ConvertTo-Html -Fragment -PreContent "<h2>No Activity Users ($noActivityCnt)</h2>"
  $f = $null
  if ($FallbackRows) {
    $f = $FallbackRows | Select-Object DisplayName,UserPrincipalName,LastSignIn,App |
         ConvertTo-Html -Fragment -PreContent "<h2>Fallback: Sign-ins ($fallbackCount)</h2>"
  }

  ConvertTo-Html -Head $style -PreContent "<h1>$Title</h1>$summary" -PostContent ($a+$n+$f) -Title $Title
}

function Parse-OneDriveUsageCsv {
  param([byte[]]$Bytes)
  $tmp = [IO.Path]::GetTempFileName()
  [IO.File]::WriteAllBytes($tmp,$Bytes)
  try { Import-Csv $tmp } finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
}

# ------------------ Primary: OneDrive usage report ------------------
$usedPrimary = $false
$reportRows = @()
if (-not $UseFallbackOnly) {
  try {
    $bytes = Get-MgReportOneDriveUsageAccountDetail -ProgressAction SilentlyContinue -Period $Period 
    $reportRows = Parse-OneDriveUsageCsv -Bytes $bytes
    if ($reportRows) { $usedPrimary = $true }
  } catch {
    Write-Warning "Primary method unavailable: $($_.Exception.Message)"
  }
}

# ------------------ Enabled users map (skip blocked/disabled) ------------------
$enabledMap = @{}
$enabledUsers = Get-MgUser -All -Property "userPrincipalName,accountEnabled,displayName" | Where-Object { $_.AccountEnabled -eq $true }
foreach ($u in $enabledUsers) { 
  if ($u.UserPrincipalName) { $enabledMap[$u.UserPrincipalName.ToLower()] = $u.DisplayName }
}

# ------------------ Lists (initialize separately; fix for PS 5.1) ------------------
$active = @()
$noAct  = @()

# ------------------ Process usage report ------------------
if ($usedPrimary) {
  foreach ($r in $reportRows) {
    $upnRaw = ($r.'User Principal Name' -as [string])
    if (-not $upnRaw) { continue }
    $upn = $upnRaw.Trim().ToLower()
    if (-not $enabledMap.ContainsKey($upn)) { continue }   # skip blocked users

    $dname = $r.'Display Name'
    $odUrl = $r.'OneDrive Url'
    $last  = $r.'Last Activity Date'

    if ($last) {
      $active += [pscustomobject]@{
        DisplayName       = $dname
        UserPrincipalName = $upnRaw.Trim()
        LastActivityDate  = (Get-Date $last).ToString('yyyy-MM-dd')
        OneDriveUrl       = $odUrl
      }
    } else {
      $noAct += [pscustomobject]@{
        DisplayName       = $dname
        UserPrincipalName = $upnRaw.Trim()
        OneDriveUrl       = $odUrl
      }
    }
  }
  $active = $active | Sort-Object DisplayName, UserPrincipalName
  $noAct  = $noAct  | Sort-Object DisplayName, UserPrincipalName
}

# ------------------ Fallback: sign-in logs filtered to OneDrive/SharePoint ------------------
$fallbackRows = $null
if (-not $usedPrimary) {
  try {
    $lookDays = switch ($Period){'D7'{7};'D30'{30};'D90'{90};'D180'{180}; default {30}}
    $sinceIso = (Get-Date).AddDays(-$lookDays).ToString('o')
    $apps = @("Office 365 SharePoint Online","Microsoft OneDrive","OneDrive Sync Engine","OneDrive Web")
    $appFilters = $apps | ForEach-Object { "appDisplayName eq '$_'" }
    $filter = "(createdDateTime ge $sinceIso) and (" + ($appFilters -join " or ") + ")"

    # Pull all sign-ins in window; we'll dedupe to latest per UPN and skip blocked
    $rows = Get-MgAuditLogSignIn -Filter $filter -All
    $latest = @{}
    foreach ($s in $rows) {
      $upn = $s.UserPrincipalName
      if ([string]::IsNullOrWhiteSpace($upn)) { continue }
      if (-not $enabledMap.ContainsKey($upn.ToLower())) { continue } # skip blocked
      $dt = $null
      try { $dt = [datetime]$s.CreatedDateTime } catch { }
      if (-not $dt) { continue }

      if (-not $latest.ContainsKey($upn) -or $dt -gt $latest[$upn].When){
        $latest[$upn] = [pscustomobject]@{
          DisplayName       = $s.UserDisplayName
          UserPrincipalName = $upn
          LastSignIn        = $dt.ToString("yyyy-MM-dd HH:mm:ss")
          App               = $s.AppDisplayName
          When              = $dt
        }
      }
    }
    # Remove helper 'When' before output
    $fallbackRows = $latest.Values | ForEach-Object {
      [pscustomobject]@{
        DisplayName       = $_.DisplayName
        UserPrincipalName = $_.UserPrincipalName
        LastSignIn        = $_.LastSignIn
        App               = $_.App
      }
    } | Sort-Object DisplayName, UserPrincipalName
  } catch {
    Write-Error "Fallback failed: $($_.Exception.Message)"
  }
}

# ------------------ Output ------------------
$html = New-ReportHtml -ActiveRows $active -NoActivityRows $noAct -FallbackRows $fallbackRows -Title "OneDrive Sign-In Report (Enabled Users Only)"
$html | Out-File -FilePath $OutputPath -Encoding UTF8

$fbCount = if ($fallbackRows) { $fallbackRows.Count } else { 0 }
Write-Host "Report written to $OutputPath"
Write-Host "Active users: $($active.Count), No activity: $($noAct.Count), Fallback: $fbCount"
