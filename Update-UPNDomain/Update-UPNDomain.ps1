[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [string]$FromDomain = 'coraswellness.onmicrosoft.com',
  [string]$ToDomain   = 'coraswellness.org',
  [switch]$UpdateMailbox
)


$ErrorActionPreference = 'Stop'

function Ensure-Module {
  param([Parameter(Mandatory)][string]$Name)
  if (-not (Get-Module -ListAvailable -Name $Name)) {
    Write-Verbose "Installing module $Name..."
    Install-Module $Name -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module $Name -Force
}

function Connect-GraphUsers {
  Ensure-Module -Name Microsoft.Graph.Users
  if (-not (Get-MgContext)) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes 'User.ReadWrite.All' | Out-Null
  }
  # Use v1.0 if available; otherwise continue on default profile
  $selectCmd = Get-Command -Name Select-MgProfile -ErrorAction SilentlyContinue
  if ($selectCmd) {
    try { Select-MgProfile -Name 'v1.0' } catch { }
  } else {
    Write-Verbose "Select-MgProfile not found; continuing with default Graph profile."
  }
}

function Connect-ExchangeIfNeeded {
  if ($UpdateMailbox) {
    Ensure-Module -Name ExchangeOnlineManagement
    if (-not (Get-ConnectionInformation)) {
      Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
      Connect-ExchangeOnline -ShowBanner:$false | Out-Null
    }
  }
}

$startTime = Get-Date
$stamp     = $startTime.ToString('yyyyMMdd-HHmmss')
$logPath   = Join-Path -Path (Get-Location) -ChildPath "UpnRename-$stamp.csv"

$results = New-Object System.Collections.Generic.List[object]

$fromPattern = [regex]::Escape("@$FromDomain") + '$'
$toSuffix    = "@$ToDomain"

Write-Host "NOTE: Ensure target domain '$ToDomain' is added/verified in your tenant." -ForegroundColor Yellow

Connect-GraphUsers
Connect-ExchangeIfNeeded

Write-Host "Querying users..." -ForegroundColor Cyan
$selectProps = @('id','displayName','userPrincipalName','onPremisesSyncEnabled','mail','accountEnabled','userType')
$users = Get-MgUser -All -Property $selectProps |
         Where-Object {
           $_.UserPrincipalName -match $fromPattern -and
           ($_.UserType -eq $null -or $_.UserType -eq 'Member')
         }

if (-not $users) {
  Write-Host "No users found with UPN ending in @$FromDomain." -ForegroundColor Yellow
  return
}

Write-Host ("Found {0} user(s) to evaluate." -f $users.Count) -ForegroundColor Green

function Test-UpnAvailable {
  param([string]$Upn)
  try {
    $found = Get-MgUser -Filter "userPrincipalName eq '$Upn'" -ConsistencyLevel eventual -CountVariable c -ErrorAction Stop
    return -not $found
  } catch {
    $fallback = Get-MgUser -All | Where-Object { $_.UserPrincipalName -eq $Upn }
    return -not $fallback
  }
}

function Set-UserUpn {
  param([Parameter(Mandatory)][object]$User, [Parameter(Mandatory)][string]$NewUpn)
  if ($PSCmdlet.ShouldProcess("$($User.UserPrincipalName) -> $NewUpn","Update-MgUser UPN")) {
    Update-MgUser -UserId $User.Id -UserPrincipalName $NewUpn
  }
}

function Update-MailboxAddresses {
  param([Parameter(Mandatory)][string]$OldUpn, [Parameter(Mandatory)][string]$NewUpn)

  $mbx = Get-Mailbox -Identity $OldUpn -ErrorAction SilentlyContinue
  if (-not $mbx) { $mbx = Get-Mailbox -Identity $NewUpn -ErrorAction SilentlyContinue }
  if (-not $mbx) { return 'NoMailbox' }

  $newPrimary = $mbx.PrimarySmtpAddress.ToString() -replace $fromPattern, $toSuffix
  $newAliases = foreach ($addr in $mbx.EmailAddresses) {
    $s = $addr.ToString()
    if ($s -like 'smtp:*' -or $s -like 'SMTP:*') { $s -replace $fromPattern, $toSuffix } else { $s }
  }

  # enforce exactly one primary "SMTP:"
  $newAliases = $newAliases | Where-Object { $_ -notlike 'SMTP:*' }
  $finalEmails = @("SMTP:$newPrimary") + ($newAliases | Sort-Object -Unique)

  if ($PSCmdlet.ShouldProcess("$OldUpn / $NewUpn","Set-Mailbox PrimarySmtp + Aliases")) {
    Set-Mailbox -Identity $mbx.Identity -PrimarySmtpAddress $newPrimary -EmailAddresses $finalEmails -EmailAddressPolicyEnabled:$false
  }
  return "MailboxUpdated:$newPrimary"
}

foreach ($u in $users) {
  $oldUpn = $u.UserPrincipalName
  $newUpn = $oldUpn -replace $fromPattern, $toSuffix
  $note   = ''
  $status = 'Skipped'

  try {
    if ($u.OnPremisesSyncEnabled) {
      $status = 'Skipped'; $note = 'On-prem synced; change in AD.'
    }
    elseif ($oldUpn -ieq $newUpn) {
      $status = 'Skipped'; $note = 'Already on target domain.'
    }
    else {
      if (-not (Test-UpnAvailable -Upn $newUpn)) {
        throw "Target UPN already exists: $newUpn"
      }
      Set-UserUpn -User $u -NewUpn $newUpn
      $status = 'UPNUpdated'; $note = 'UPN changed via Graph.'
      if ($UpdateMailbox) { $note = "$note $(Update-MailboxAddresses -OldUpn $oldUpn -NewUpn $newUpn)" }
    }
  }
  catch {
    $status = if ($status -eq 'UPNUpdated') { 'UPNUpdatedWithErrors' } else { 'Failed' }
    $note   = if ($note) { "$note | $($_.Exception.Message)" } else { $_.Exception.Message }
  }
  finally {
    $results.Add([pscustomobject]@{
      TimeStamp      = (Get-Date).ToString('s')
      DisplayName    = $u.DisplayName
      OldUPN         = $oldUpn
      NewUPN         = $newUpn
      AccountEnabled = $u.AccountEnabled
      OnPremSynced   = [bool]$u.OnPremisesSyncEnabled
      Status         = $status
      Notes          = $note
    })
  }
}

$results | Tee-Object -Variable out | Format-Table -AutoSize
$results | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $logPath
Write-Host "`nComplete. Results exported to: $logPath" -ForegroundColor Green
if ($WhatIf) { Write-Host "NOTE: -WhatIf was used. No changes were committed." -ForegroundColor Yellow }
