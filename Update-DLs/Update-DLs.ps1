<# 
.SYNOPSIS
  Enable external delivery for all distribution lists in Exchange Online.

.DESCRIPTION
  Sets RequireSenderAuthenticationEnabled:$false on all distribution groups
  (includes mail-enabled universal distribution and security groups).
  Runs in Preview mode by default (no changes). Use -Apply to make changes.

.PARAMETER Apply
  Actually apply the changes (otherwise runs with -WhatIf).

.PARAMETER ReportPath
  Path for CSV report (before/after view). Default: .\DL-ExternalDelivery-Report.csv

.EXAMPLE
  .\Enable-ExternalDelivery-AllDLs.ps1              # Preview only
  .\Enable-ExternalDelivery-AllDLs.ps1 -Apply       # Make changes
  .\Enable-ExternalDelivery-AllDLs.ps1 -Apply -ReportPath C:\temp\report.csv
#>

[CmdletBinding()]
param(
  [switch]$Apply,
  [string]$ReportPath = ".\DL-ExternalDelivery-Report.csv"
)

function Ensure-Module {
  param([string]$Name)
  if (-not (Get-Module -ListAvailable -Name $Name)) {
    Write-Host "Installing module '$Name' for current user..." -ForegroundColor Yellow
    Install-Module $Name -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module $Name -Force
}

try {
  Ensure-Module -Name ExchangeOnlineManagement

  # Interactive sign-in; swap to App-only if you use app certs in your environment.
  Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
  Connect-ExchangeOnline -ShowBanner:$false

  Write-Host "Querying distribution groups..." -ForegroundColor Cyan
  $groups = Get-DistributionGroup -ResultSize Unlimited | Sort-Object DisplayName

  if (-not $groups) {
    Write-Host "No distribution groups found." -ForegroundColor Yellow
    return
  }

  $pre = $groups | Select-Object DisplayName, PrimarySmtpAddress, Identity, GroupType,
                         RecipientTypeDetails, RequireSenderAuthenticationEnabled

  $whatIf = $false
  if ($Apply.IsPresent) { $whatIf = $false }

  $changed = @()
  foreach ($g in $groups) {
    if ($g.RequireSenderAuthenticationEnabled) {
      Write-Host ("Enabling external delivery for: {0} <{1}>" -f $g.DisplayName, $g.PrimarySmtpAddress) `
        -ForegroundColor Green
      Set-DistributionGroup -Identity $g.Identity -RequireSenderAuthenticationEnabled:$false -WhatIf:$whatIf
      if (-not $whatIf) { $changed += $g.Identity }
    } else {
      Write-Host ("Already allows external senders: {0}" -f $g.DisplayName) -ForegroundColor DarkGray
    }
  }

  # Re-query to capture the post-change state (or simulated, if Preview)
  $post = Get-DistributionGroup -ResultSize Unlimited | Select-Object `
            DisplayName, PrimarySmtpAddress, Identity, GroupType,
            RecipientTypeDetails, RequireSenderAuthenticationEnabled

  # Build a comparison report
  $report = foreach ($p in $post) {
    $before = $pre | Where-Object { $_.Identity -eq $p.Identity } | Select-Object -First 1
    [pscustomobject]@{
      DisplayName                         = $p.DisplayName
      PrimarySmtpAddress                  = $p.PrimarySmtpAddress
      Identity                            = $p.Identity
      RecipientTypeDetails                = $p.RecipientTypeDetails
      Before_RequireSenderAuthEnabled     = if ($before) { $before.RequireSenderAuthenticationEnabled } else { $null }
      After_RequireSenderAuthEnabled      = $p.RequireSenderAuthenticationEnabled
      ChangedInThisRun                    = ($changed -contains $p.Identity)
      Mode                                = if ($whatIf) { "Preview (WhatIf)" } else { "Applied" }
      TimestampUtc                        = (Get-Date).ToUniversalTime()
    }
  }

  $report | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $ReportPath
  Write-Host "Report written to: $ReportPath" -ForegroundColor Cyan

} catch {
  Write-Error $_
} finally {
  Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
}
