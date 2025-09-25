[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [string]$DefaultLocation,
  [string]$ResultsPath
)

# Ensure Microsoft Graph Users module
if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) {
  Write-Host "Microsoft.Graph module not found. Installing for current user..." -ForegroundColor Yellow
  Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph.Users -ErrorAction Stop

# Connect to Graph
$Scopes = @('User.ReadWrite.All')
try {
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $Scopes }
} catch {
  Connect-MgGraph -Scopes $Scopes
}

# Pull users
Write-Host "Retrieving users from Microsoft Graph..." -ForegroundColor Cyan
$users = Get-MgUser -All -Property Id,UserPrincipalName,DisplayName,OfficeLocation

$targets = $users |
  Where-Object { [string]::IsNullOrWhiteSpace($_.OfficeLocation) } |
  Sort-Object UserPrincipalName

if (-not $targets) {
  Write-Host "All users already have an OfficeLocation set. ✅" -ForegroundColor Green
  return
}

Write-Host ("Found {0} user(s) with blank OfficeLocation." -f $targets.Count) -ForegroundColor Yellow

function Choose-Action {
  param(
    [string]$Upn,
    [string]$DisplayName,
    [string]$LastLocation
  )
  Write-Host ""
  Write-Host "User: $DisplayName <$Upn> (OfficeLocation is blank)" -ForegroundColor Cyan
  if ($LastLocation) { Write-Host "Last location entered: $LastLocation" }
  Write-Host "[E]nter   [R]epeat-last   [S]kip   [Q]uit"
  while ($true) {
    $resp = Read-Host "Choose action (E/R/S/Q)"
    switch ($resp.ToUpper()) {
      'E' { return 'Enter' }
      'R' { return 'Repeat' }
      'S' { return 'Skip' }
      'Q' { return 'Quit' }
      default { Write-Host "Please enter E, R, S, or Q." -ForegroundColor DarkYellow }
    }
  }
}

$results = New-Object System.Collections.Generic.List[object]
$lastLocation = $null
$quit = $false

foreach ($u in $targets) {
  if ($quit) { break }

  $upn = $u.UserPrincipalName
  $dn  = if ($u.DisplayName) { $u.DisplayName } else { $upn }

  :promptUser while ($true) {
    $action = Choose-Action -Upn $upn -DisplayName $dn -LastLocation $lastLocation

    switch ($action) {
      'Enter' {
        $entered = Read-Host ("Enter OfficeLocation for {0} [{1}]" -f $dn, $(if($DefaultLocation){ "default: $DefaultLocation" } else { "no default" }))
        if ([string]::IsNullOrWhiteSpace($entered)) {
          if ($DefaultLocation) { $entered = $DefaultLocation } else {
            Write-Host "No value provided; returning to options." -ForegroundColor DarkYellow
            continue promptUser
          }
        }

        if ($WhatIf) {
          Write-Host ("[WhatIf] Would set OfficeLocation for {0} -> '{1}'" -f $upn, $entered)
          $results.Add([pscustomobject]@{
            Timestamp   = Get-Date
            UPN         = $upn
            DisplayName = $dn
            OldLocation = $null
            NewLocation = $entered
            Action      = 'WhatIf-Update'
            Result      = 'No change (dry run)'
          })
          $lastLocation = $entered
          break promptUser
        }

        try {
          if ($PSCmdlet.ShouldProcess($upn, "Update OfficeLocation -> '$entered'")) {
            Update-MgUser -UserId $u.Id -OfficeLocation $entered
            $post = Get-MgUser -UserId $u.Id -Property OfficeLocation
            $ok = ($post.OfficeLocation -eq $entered)
            $results.Add([pscustomobject]@{
              Timestamp   = Get-Date
              UPN         = $upn
              DisplayName = $dn
              OldLocation = $null
              NewLocation = $entered
              Action      = 'Updated'
              Result      = $(if($ok){'Success'} else {"Updated; verify (now: '$($post.OfficeLocation)')"})
            })
            $lastLocation = $entered
          }
        } catch {
          Write-Warning "Failed to update $upn $($_.Exception.Message)"
          $results.Add([pscustomobject]@{
            Timestamp   = Get-Date
            UPN         = $upn
            DisplayName = $dn
            OldLocation = $null
            NewLocation = $entered
            Action      = 'Update'
            Result      = "Failed: $($_.Exception.Message)"
          })
        }
        break promptUser
      }

      'Repeat' {
        if ([string]::IsNullOrWhiteSpace($lastLocation)) {
          Write-Host "No previous location to repeat." -ForegroundColor DarkYellow
          continue promptUser
        }

        if ($WhatIf) {
          Write-Host ("[WhatIf] Would set OfficeLocation for {0} -> '{1}'" -f $upn, $lastLocation)
          $results.Add([pscustomobject]@{
            Timestamp   = Get-Date
            UPN         = $upn
            DisplayName = $dn
            OldLocation = $null
            NewLocation = $lastLocation
            Action      = 'WhatIf-Update'
            Result      = 'No change (dry run)'
          })
          break promptUser
        }

        try {
          if ($PSCmdlet.ShouldProcess($upn, "Update OfficeLocation -> '$lastLocation'")) {
            Update-MgUser -UserId $u.Id -OfficeLocation $lastLocation
            $post = Get-MgUser -UserId $u.Id -Property OfficeLocation
            $ok = ($post.OfficeLocation -eq $lastLocation)
            $results.Add([pscustomobject]@{
              Timestamp   = Get-Date
              UPN         = $upn
              DisplayName = $dn
              OldLocation = $null
              NewLocation = $lastLocation
              Action      = 'Updated'
              Result      = $(if($ok){'Success'} else {"Updated; verify (now: '$($post.OfficeLocation)')"})
            })
          }
        } catch {
          Write-Warning "Failed to update $upn $($_.Exception.Message)"
          $results.Add([pscustomobject]@{
            Timestamp   = Get-Date
            UPN         = $upn
            DisplayName = $dn
            OldLocation = $null
            NewLocation = $lastLocation
            Action      = 'Update'
            Result      = "Failed: $($_.Exception.Message)"
          })
        }
        break promptUser
      }

      'Skip' {
        $results.Add([pscustomobject]@{
          Timestamp   = Get-Date
          UPN         = $upn
          DisplayName = $dn
          OldLocation = $null
          NewLocation = $null
          Action      = 'Skipped'
          Result      = 'No change'
        })
        break promptUser
      }

      'Quit' {
        Write-Host "Quitting at your request." -ForegroundColor Cyan
        $quit = $true
        break promptUser
      }
    }
  }
}

# Export results
if (-not $ResultsPath) {
  $scriptDir = Split-Path -Parent $PSCommandPath
  if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $ResultsPath = Join-Path $scriptDir "Fill-OfficeLocation-Interactive_$stamp.csv"
}

$results | Export-Csv -Path $ResultsPath -NoTypeInformation -Encoding UTF8
Write-Host "`nDone. Results saved to: $ResultsPath" -ForegroundColor Green
