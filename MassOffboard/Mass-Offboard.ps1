<#
.SYNOPSIS
    Mass offboarding helper for Microsoft 365 / Exchange Online users.

.DESCRIPTION
    This script automates common offboarding actions for a list of user accounts
    in a Microsoft 365 tenant.

    For each UPN in the specified input file, the script:
      - Looks up the user in Microsoft Graph.
      - Converts the user's Exchange Online mailbox to a Shared mailbox.
      - Removes all *directly assigned* licenses from the user (group-based
        licenses are left in place, but called out in the log).
      - Blocks the user's sign-in in Entra ID (accountEnabled = $false).
      - (Optional) Revokes sign-in sessions via Microsoft Graph if the
        Revoke-SignInSessions helper is present in the script.

    All actions are logged to a CSV file with a timestamped filename by default.
    Failures in individual steps are captured per-user so that one bad account
    does not stop processing for the remaining users.

.PARAMETER UserListPath
    Path to a text file containing one user principal name (UPN) per line.
    Lines starting with '#' or blank lines are ignored.
    Default: .\Users.txt

.PARAMETER LogPath
    Path to the CSV log file to create.
    By default, a file named "SharedMailbox_Conversion_Log_yyyyMMdd_HHmmss.csv"
    is created in the current directory.

.INPUTS
    None. You cannot pipe input directly to this script.

.OUTPUTS
    None. The script writes status output to the console and a CSV log file
    with one row per processed user.

.EXAMPLE
    .\Mass-Offboard.ps1

    Uses the default Users.txt file in the current directory and writes a
    timestamped CSV log file with results of the mailbox conversion, license
    removal, and sign-in blocking actions.

.EXAMPLE
    .\Mass-Offboard.ps1 -UserListPath .\Users.txt

    Processes the list of UPNs in Users.txt instead of the default
    Users.txt and writes a timestamped CSV log of actions taken.

.EXAMPLE
    .\Mass-Offboard.ps1 -UserListPath .\Users.txt -LogPath .\Logs\Offboard_Log.csv

    Processes the specified list of users and writes the results to
    .\Logs\Offboard_Log.csv.

.NOTES
    Author:   Dan Nelson / dnelson@reduxtc.com 
    Requires: ExchangeOnlineManagement module
              Microsoft.Graph.Authentication
              Microsoft.Graph.Users
              Microsoft.Graph.Users.Actions

    Permissions:
      - The account running this script must be able to:
          * Connect to Exchange Online and manage mailboxes.
          * Connect to Microsoft Graph with at least:
              User.ReadWrite.All
              Directory.ReadWrite.All
          * (Optional for session revocation)
              Directory.AccessAsUser.All or equivalent permissions for
              POST /users/{id}/revokeSignInSessions.

    Run this script from an elevated PowerShell session if module installation
    is required for the current user.
#>


[CmdletBinding()]
param(
    [string]$UserListPath = ".\Users.txt",
    [string]$LogPath = ".\SharedMailbox_Conversion_Log_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date)
)

$GraphMinVersion = '2.29.0'
# --- Helpers --- #
function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name,[string]$MinimumVersion)
    $installed = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    $needInstall = -not $installed
    if ($installed -and $MinimumVersion) { $needInstall = ([version]$installed.Version -lt [version]$MinimumVersion) }
    if ($needInstall) {
        Write-Host "Installing module $Name (>= $MinimumVersion)..." -ForegroundColor Yellow
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber -MinimumVersion $MinimumVersion -ErrorAction Stop
    }
    Import-Module $Name -ErrorAction Stop
}

function Ensure-CloudModules {
    # Exchange Online
    Ensure-Module -Name ExchangeOnlineManagement
    # Graph (only what we need... just the tip)
    Ensure-Module -Name Microsoft.Graph.Authentication -MinimumVersion $GraphMinVersion
    Ensure-Module -Name Microsoft.Graph.Users          -MinimumVersion $GraphMinVersion
    Ensure-Module -Name Microsoft.Graph.Users.Actions  -MinimumVersion $GraphMinVersion
}

function Connect-Cloud {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false

    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    $scopes = @('User.ReadWrite.All','Directory.ReadWrite.All')
    Connect-MgGraph -Scopes $scopes -NoWelcome
    }

function Get-UserIdOrThrow {
    param([string]$Upn)
    try {
        Get-MgUser -UserId $Upn -Property "id,userPrincipalName,accountEnabled,assignedLicenses" -ErrorAction Stop
    } catch {
        throw "Graph lookup failed for '$Upn': $($_.Exception.Message)"
    }
}

function Convert-ToSharedMailbox { # because who wants to do this by hand? 
    param([string]$Identity)
    try {
        Start-Sleep -Seconds 10
        $mbx = Get-ExoMailbox -Identity $Identity -ErrorAction Stop

        if ($mbx.RecipientTypeDetails -eq 'SharedMailbox') {
            return "Already Shared"
        }

        Set-Mailbox -Identity $Identity -Type Shared -ErrorAction Stop

        # Wait up to ~60 seconds for the type to flip... I DO NOT CARE HOW LONG THIS MAKES IT TAKE 
        $maxAttempts = 12 # you can change this if you want, but this is a generally good value 
        for ($i = 1; $i -le $maxAttempts; $i++) {
            Start-Sleep -Seconds 5
            $mbx2 = Get-ExoMailbox -Identity $Identity -ErrorAction Stop
            if ($mbx2.RecipientTypeDetails -eq 'SharedMailbox') {
                return "Converted to Shared"
            }
        }

        # If we get here, the command succeeded but the type hasn't updated yet, because replication is a gigantic fucking asshole 
        return "Conversion requested; mailbox not yet reporting SharedMailbox (likely replication delay)"
    } catch {
        throw "Mailbox conversion error for '$Identity': $($_.Exception.Message)"
    }
}

function Remove-AllLicenses {
    param([string]$UserId, [object]$UserObject)
    try {
        $skuIds = @()
        if ($UserObject.AssignedLicenses -and $UserObject.AssignedLicenses.Count -gt 0) {
            $skuIds = $UserObject.AssignedLicenses.SkuId
        } else {
            $details = Get-MgUserLicenseDetail -UserId $UserId -ErrorAction SilentlyContinue
            if ($details) { $skuIds = $details.SkuId }
        }

        if (-not $skuIds -or $skuIds.Count -eq 0) { return "No licenses assigned" }

        # Remove direct licenses (leave AddLicenses empty)
        Set-MgUserLicense -UserId $UserId -RemoveLicenses $skuIds -AddLicenses @() -ErrorAction Stop

        Start-Sleep -Seconds 3
        $check = Get-MgUser -UserId $UserId -Property "assignedLicenses" -ErrorAction Stop
        if ($check.AssignedLicenses.Count -eq 0) { 
            "Removed all licenses"
        } else {
            "Some licenses remain (likely group-based). Review group memberships."
        }
    } catch {
        throw "License removal error: $($_.Exception.Message)"
    }
}

function Block-SignIn {
    param([string]$UserId)
    try {
        Update-MgUser -UserId $UserId -AccountEnabled:$false -ErrorAction Stop
        $verify = Get-MgUser -UserId $UserId -Property "accountEnabled" -ErrorAction Stop
        if ($verify.AccountEnabled -eq $false) { "Blocked sign-in" } else { throw "Account still enabled after update." }
    } catch {
        throw "Block sign-in error: $($_.Exception.Message)"
    }
}

function Revoke-SignInSessions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,   # GUID, not UPN
        [Parameter(Mandatory = $false)]
        [string]$Upn       # purely for logging and because IT'S AWESOME!!!!!!!! i don't know i'm sleep deprived 
    )
    try {
        # This calls POST /users/{id}/revokeSignInSessions
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$UserId/revokeSignInSessions" -ErrorAction Stop
        if ($Upn) {
            return "Revoke sign-in sessions requested for $Upn"
        } else {
            return "Revoke sign-in sessions requested"
        }
    }
    catch {
        throw "Failed to revoke sign-in sessions for '$Upn' ($UserId) $($_.Exception.Message)"
    }
}


# --- Main ---
if (-not (Test-Path -Path $UserListPath)) { throw "User list not found at '$UserListPath'." }

Ensure-CloudModules
Connect-Cloud

$log = New-Object System.Collections.Generic.List[Object]
$upns = Get-Content -Path $UserListPath | ForEach-Object { $_.Trim() } | Where-Object { $_ -and ($_ -notmatch '^\s*#') }
if (-not $upns -or $upns.Count -eq 0) { throw "No UPNs found in '$UserListPath'." }

Write-Host "`nProcessing $($upns.Count) user(s)..." -ForegroundColor Green

foreach ($upn in $upns) {
    Write-Host "---- $upn ----" -ForegroundColor White
    $result = [ordered]@{
        Timestamp          = (Get-Date).ToString("s")
        UPN                = $upn
        ConvertMailbox     = $null
        LicenseAction      = $null
        SignInAction       = $null
		SessionAction      = $null   # <-- new 11-19-25
        Notes              = $null
        Status             = "Success"
    }

    try {
        $user = Get-UserIdOrThrow -Upn $upn

        try { $result.ConvertMailbox = Convert-ToSharedMailbox -Identity $upn } catch { $result.ConvertMailbox = "Failed"; throw }
        try { $result.LicenseAction  = Remove-AllLicenses -UserId $user.Id -UserObject $user } catch { $result.LicenseAction = "Failed"; throw }
		try { $result.SessionAction = Revoke-SignInSessions -UserId $user.Id -Upn $upn } catch { $result.SessionAction = "Failed: $($_.Exception.Message)" }
        try { $result.SignInAction   = Block-SignIn -UserId $user.Id } catch { $result.SignInAction = "Failed"; throw }

    } catch {
        $result.Status = "Error"
        $result.Notes  = $_.Exception.Message
        Write-Warning "Error for $upn $($result.Notes)"
    }

    $log.Add([pscustomobject]$result)
}

$log | Tee-Object -Variable OutLog | Format-Table -AutoSize
$OutLog | Export-Csv -NoTypeInformation -Path $LogPath -Encoding UTF8
Write-Host "`nLog written to: $LogPath" -ForegroundColor Green

try { Disconnect-ExchangeOnline -Confirm:$false } catch {}
try { Disconnect-MgGraph -Confirm:$false } catch {
	#You didn't really expect anything to be here did you?
}