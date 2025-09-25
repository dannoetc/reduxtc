<# Minimal-Graph; lookup via -Filter (not -UserId) to avoid GUID-only error #>

$ErrorActionPreference = 'Stop'

$InputFile = Join-Path (Get-Location) 'unfiltered.txt'
$CsvOut    = Join-Path (Get-Location) 'SignInActivityReport.csv'
$NeverOut  = Join-Path (Get-Location) 'users.txt'

if (-not (Test-Path $InputFile)) { throw "Input file not found: $InputFile" }

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

# Connect (adjust scopes to your consent model)
$scopes = @('User.Read.All','AuditLog.Read.All')
Connect-MgGraph -Scopes $scopes | Out-Null

# Load UPNs
$upns = Get-Content -LiteralPath $InputFile |
    Where-Object { $_ -and $_.Trim() -ne '' -and $_ -notmatch '^\s*#' } |
    ForEach-Object { $_.Trim() } |
    Select-Object -Unique
if (-not $upns) { throw "No UPNs found in $InputFile." }

function Get-UserByUpn {
    param([Parameter(Mandatory)][string]$Upn)

    # Escape single quotes for OData filter
    $escaped = $Upn -replace "'", "''"

    # -ConsistencyLevel eventual unlocks advanced query features in Graph
    # Use -Top 1; still guard for 0/1+ results
    $u = Get-MgUser `
            -Filter "userPrincipalName eq '$escaped'" `
            -ConsistencyLevel eventual `
            -Property "id,userPrincipalName,signInActivity" `
            -Top 1

    # If nothing came back, return $null
    if (-not $u) { return $null }

    # Some tenants can return fuzzy results with -Search; we used -Filter so it should be exact,
    # but be defensive anyway.
    if ($u.userPrincipalName -ne $Upn) {
        # Try case-insensitive exact match if needed
        $u = $u | Where-Object { $_.userPrincipalName -ieq $Upn } | Select-Object -First 1
    }
    return $u
}

$results   = New-Object System.Collections.Generic.List[object]
$neverList = New-Object System.Collections.Generic.List[string]

$i = 0
$total = $upns.Count
foreach ($upn in $upns) {
    $i++
    Write-Progress -Activity "Checking sign-in activity" -Status "$i of $total" -PercentComplete (($i/$total)*100)

    $row = [ordered]@{
        UPN                      = $upn
        FoundInTenant            = $false
        LastInteractiveSignIn    = $null
        LastNonInteractiveSignIn = $null
        HasEverSignedIn          = $false
        Notes                    = ''
    }

    try {
        $user = Get-UserByUpn -Upn $upn

        if ($user) {
            $row.FoundInTenant = $true
            $sia = $user.SignInActivity
            if ($sia) {
                $row.LastInteractiveSignIn    = $sia.LastSignInDateTime
                $row.LastNonInteractiveSignIn = $sia.LastNonInteractiveSignInDateTime
                $row.HasEverSignedIn = [bool]($row.LastInteractiveSignIn -or $row.LastNonInteractiveSignIn)
                if (-not $row.HasEverSignedIn) { $row.Notes = 'No recorded sign-in timestamps.' }
            } else {
                $row.Notes = 'signInActivity is null (no activity or insufficient perms).'
            }
        } else {
            $row.Notes = 'User not found by UPN.'
        }
    }
    catch {
        $row.Notes = "Error: $($_.Exception.Message)"
    }

    $results.Add([pscustomobject]$row) | Out-Null
    if (-not $row.HasEverSignedIn) { $neverList.Add($upn) | Out-Null }

    Start-Sleep -Milliseconds 100
}

# Outputs
$results | Sort-Object UPN | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $CsvOut
$neverList | Sort-Object | Set-Content -Path $NeverOut -Encoding UTF8

Write-Host "`nDone." -ForegroundColor Green
Write-Host "CSV report: $CsvOut"
Write-Host "Never-signed-in UPNs: $NeverOut" -ForegroundColor Yellow
