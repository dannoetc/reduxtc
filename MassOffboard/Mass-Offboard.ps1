[CmdletBinding()]
param(
    [string]$UserListPath = ".\Users.txt",
    [string]$LogPath = ".\SharedMailbox_Conversion_Log_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date)
)

$GraphMinVersion = '2.29.0'

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
    # Graph (only what we need)
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
    Select-MgProfile -Name "v1.0"
}

function Get-UserIdOrThrow {
    param([string]$Upn)
    try {
        Get-MgUser -UserId $Upn -Property "id,userPrincipalName,accountEnabled,assignedLicenses" -ErrorAction Stop
    } catch {
        throw "Graph lookup failed for '$Upn': $($_.Exception.Message)"
    }
}

function Convert-ToSharedMailbox {
    param([string]$Identity)
    try {
        Start-Sleep -Seconds 10
        $mbx = Get-ExoMailbox -Identity $Identity -ErrorAction Stop

        if ($mbx.RecipientTypeDetails -eq 'SharedMailbox') {
            return "Already Shared"
        }

        Set-Mailbox -Identity $Identity -Type Shared -ErrorAction Stop

        # Wait up to ~60 seconds for the type to flip
        $maxAttempts = 12
        for ($i = 1; $i -le $maxAttempts; $i++) {
            Start-Sleep -Seconds 5
            $mbx2 = Get-ExoMailbox -Identity $Identity -ErrorAction Stop
            if ($mbx2.RecipientTypeDetails -eq 'SharedMailbox') {
                return "Converted to Shared"
            }
        }

        # If we get here, the command succeeded but the type hasn't updated yet
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
        Notes              = $null
        Status             = "Success"
    }

    try {
        $user = Get-UserIdOrThrow -Upn $upn

        try { $result.ConvertMailbox = Convert-ToSharedMailbox -Identity $upn } catch { $result.ConvertMailbox = "Failed"; throw }
        try { $result.LicenseAction  = Remove-AllLicenses -UserId $user.Id -UserObject $user } catch { $result.LicenseAction = "Failed"; throw }
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
