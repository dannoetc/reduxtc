<#
.SYNOPSIS
  Tenant-wide sign-in activity report via Microsoft Graph (minimal modules).
  Exports full CSV + text file of users who HAVE signed in.

.OUTPUT
  - TenantSignInActivityReport.csv   (all users, with details)
  - signedinusers.txt                (only UPNs of users that have signed in)
#>

$ErrorActionPreference = 'Stop'

$CsvOut   = Join-Path (Get-Location) 'TenantSignInActivityReport.csv'
$SignedIn = Join-Path (Get-Location) 'signedinusers.txt'

# --- Minimal modules only ---
$requiredModules = @('Microsoft.Graph.Authentication','Microsoft.Graph.Users')
foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "Installing $m..." -ForegroundColor Yellow
        Install-Module $m -Scope AllUsers -Force
    }
    Import-Module $m -ErrorAction Stop
}

try { Select-MgProfile -Name 'v1.0' } catch {}

$scopes = @('User.Read.All','AuditLog.Read.All')
Connect-MgGraph -Scopes $scopes | Out-Null

# Helper for transient Graph errors
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)] [scriptblock] $Script,
        [int] $MaxRetries = 5,
        [int] $BaseDelayMs = 500
    )
    $attempt = 0
    while ($true) {
        try {
            return & $Script
        } catch {
            $attempt++
            $msg = $_.Exception.Message
            if ($attempt -ge $MaxRetries -or ($msg -notmatch '(^|\s)(429|5\d{2})')) { throw }
            $delay = [math]::Min(15000, $BaseDelayMs * [math]::Pow(2, ($attempt - 1)))
            Write-Warning "Transient error (attempt $attempt/$MaxRetries): $msg`nRetrying in $([int]$delay)ms…"
            Start-Sleep -Milliseconds $delay
        }
    }
}

Write-Host "Querying users..." -ForegroundColor Cyan

$select = 'id,userPrincipalName,displayName,accountEnabled,signInActivity'
$users  = Invoke-WithRetry { 
    Get-MgUser -All -Property $select -ConsistencyLevel eventual 
}

$results       = New-Object System.Collections.Generic.List[object]
$signedInUsers = New-Object System.Collections.Generic.List[string]

$total = ($users | Measure-Object).Count
$index = 0

foreach ($u in $users) {
    $index++
    if ($index % 50 -eq 0) {
        Write-Progress -Activity "Processing users" -Status "$index of $total" -PercentComplete (($index / [math]::Max(1,$total)) * 100)
    }

    $sia = $u.SignInActivity

    $lastInteractive    = $null
    $lastNonInteractive = $null
    $hasEver            = $false
    $notes              = ''

    if ($sia) {
        $lastInteractive    = $sia.LastSignInDateTime
        $lastNonInteractive = $sia.LastNonInteractiveSignInDateTime
        $hasEver            = [bool]($lastInteractive -or $lastNonInteractive)
        if (-not $hasEver) { $notes = 'No recorded sign-in timestamps.' }
    } else {
        $notes = 'signInActivity is null (no activity or insufficient perms).'
    }

    $results.Add([pscustomobject][ordered]@{
        UPN                      = $u.UserPrincipalName
        DisplayName              = $u.DisplayName
        AccountEnabled           = $u.AccountEnabled
        LastInteractiveSignIn    = $lastInteractive
        LastNonInteractiveSignIn = $lastNonInteractive
        HasEverSignedIn          = $hasEver
        Notes                    = $notes
    }) | Out-Null

    if ($hasEver) { $signedInUsers.Add($u.UserPrincipalName) | Out-Null }
}

# Export CSV (all users)
$results | Sort-Object UPN | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $CsvOut

# Export only signed-in users (UPNs)
$signedInUsers | Sort-Object | Set-Content -Path $SignedIn -Encoding UTF8

Write-Host "`nDone." -ForegroundColor Green
Write-Host "CSV report: $CsvOut"
Write-Host "Signed-in UPNs: $SignedIn" -ForegroundColor Yellow
