<# 
.SYNOPSIS
    Convert mailboxes to Shared, remove all licenses, and block sign-in for a list of UPNs.

.DESCRIPTION
    - Reads UPNs from .\Users.txt (one per line; lines starting with # are ignored)
    - Connects to Exchange Online and Microsoft Graph
    - Converts mailbox type to Shared (if not already)
    - Removes all user-assigned licenses
    - Disables the account (blocks sign-in)
    - Writes a CSV log with per-user results and notes

.NOTES
    Run in an elevated PowerShell session with internet access for module install & auth.
#>

[CmdletBinding()]
param(
    [string]$UserListPath = ".\Users.txt",
    [string]$LogPath = ".\SharedMailbox_Conversion_Log_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date)
)

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name,[string]$MinimumVersion)
    $installed = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $installed -or ($MinimumVersion -and [version]$installed.Version -lt [version]$MinimumVersion)) {
        Write-Host "Installing module $Name..." -ForegroundColor Yellow
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module $Name -ErrorAction Stop
}

function Connect-Cloud {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false

    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    $scopes = @('User.ReadWrite.All','Directory.ReadWrite.All')
    Connect-MgGraph -Scopes $scopes -NoWelcome
    Select-MgProfile -Name "v1.0"
}

function Get-UserIdOrThrow {
    param([string]$Upn)
    try {
        # Ask Graph for the user; request AssignedLicenses up front for fewer calls
        $u = Get-MgUser -UserId $Upn -Property "id,userPrincipalName,accountEnabled,assignedLicenses" -ErrorAction Stop
        return $u
    } catch {
        throw "Graph lookup failed for '$Upn': $($_.Exception.Message)"
    }
}

function Convert-ToSharedMailbox {
    param([string]$Identity)
    try {
        $mbx = Get-ExoMailbox -Identity $Identity -ErrorAction Stop
        if ($mbx.RecipientTypeDetails -ne 'SharedMailbox') {
            Set-Mailbox -Identity $Identity -Type Shared -ErrorAction Stop
            # Refresh to confirm
            $mbx2 = Get-ExoMailbox -Identity $Identity -ErrorAction Stop
            if ($mbx2.RecipientTypeDetails -eq 'SharedMailbox') {
                return "Converted to Shared"
            } else {
                throw "Post-conversion check did not return SharedMailbox."
            }
        } else {
            return "Already Shared"
        }
    } catch {
        throw "Mailbox conversion error for '$Identity': $($_.Exception.Message)"
    }
}

function Remove-AllLicenses {
    param([string]$UserId, [object]$UserObject)
    # $UserObject is the MgUser we already fetched (contains AssignedLicenses)
    try {
        $assignedSkuIds = @()
        if ($null -ne $UserObject.AssignedLicenses -and $UserObject.AssignedLicenses.Count -gt 0) {
            $assignedSkuIds = $UserObject.AssignedLicenses.SkuId
        } else {
            # Double-check with licenseDetails (also gives skuPartNumber for logging)
            $details = Get-MgUserLicenseDetail -UserId $UserId -ErrorAction SilentlyContinue
            if ($details) { $assignedSkuIds = $details.SkuId }
        }

        if (-not $assignedSkuIds -or $assignedSkuIds.Count -eq 0) {
            return "No licenses assigned"
        }

        Update-MgUserLicense -UserId $UserId -RemoveLicenses $assignedSkuIds -AddLicenses @{} -ErrorAction Stop

        Start-Sleep -Seconds 3
        # Verify removal
        $check = Get-MgUser -UserId $UserId -Property "assignedLicenses" -ErrorAction Stop
        if ($check.AssignedLicenses.Count -eq 0) {
            return "Removed all licenses"
        } else {
            # This is often due to group-based licensing
            return "Some licenses remain (likely group-based). Review group memberships."
        }
    } catch {
        throw "License removal error: $($_.Exception.Message)"
    }
}

function Block-SignIn {
    param([string]$UserId)
    try {
        Update-MgUser -UserId $UserId -AccountEnabled:$false -ErrorAction Stop
        # Verify
        $verify = Get-MgUser -UserId $UserId -Property "accountEnabled" -ErrorAction Stop
        if ($verify.AccountEnabled -eq $false) { 
            return "Blocked sign-in"
        } else {
            throw "Account still enabled after update."
        }
    } catch {
        throw "Block sign-in error: $($_.Exception.Message)"
    }
}

# --- Main ---

# Safety checks
if (-not (Test-Path -Path $UserListPath)) {
    throw "User list not found at '$UserListPath'. Place a file with one UPN per line."
}

# Ensure modules
Ensure-Module -Name ExchangeOnlineManagement
Ensure-Module -Name Microsoft.Graph

# Connect
Connect-Cloud

# Prepare logging
$log = New-Object System.Collections.Generic.List[Object]

# Read list
$upns = Get-Content -Path $UserListPath | ForEach-Object { $_.Trim() } | Where-Object { $_ -and ($_ -notmatch '^\s*#') }

if (-not $upns -or $upns.Count -eq 0) {
    throw "No UPNs found in '$UserListPath'."
}

Write-Host "`nProcessing $($upns.Count) user(s)..." -ForegroundColor Green

foreach ($upn in $upns) {
    Write-Host "---- $upn ----" -ForegroundColor White
    $result = [ordered]@{
        Timestamp          = (Get-Date).ToString("s")
        UPN                = $upn
        ConvertMailbox     = $null
        LicenseAction      = $null
        SignInAction       = $null
        Notes              = $null
        Status             = "Success"
    }

    try {
        $user = Get-UserIdOrThrow -Upn $upn

        # Convert mailbox to Shared
        try {
            $conv = Convert-ToSharedMailbox -Identity $upn
            $result.ConvertMailbox = $conv
        } catch {
            $result.ConvertMailbox = "Failed"
            throw
        }

        # Remove all licenses
        try {
            $lic = Remove-AllLicenses -UserId $user.Id -UserObject $user
            $result.LicenseAction = $lic
        } catch {
            $result.LicenseAction = "Failed"
            throw
        }

        # Block sign-in
        try {
            $block = Block-SignIn -UserId $user.Id
            $result.SignInAction = $block
        } catch {
            $result.SignInAction = "Failed"
            throw
        }
    } catch {
        $result.Status = "Error"
        $result.Notes  = $_.Exception.Message
        Write-Warning "Error for $upn: $($result.Notes)"
    }

    $log.Add([pscustomobject]$result)
}

# Output & save log
$log | Tee-Object -Variable OutLog | Format-Table -AutoSize
$OutLog | Export-Csv -NoTypeInformation -Path $LogPath -Encoding UTF8

Write-Host "`nLog written to: $LogPath" -ForegroundColor Green

# Disconnect cleanly
try { Disconnect-ExchangeOnline -Confirm:$false } catch {}
