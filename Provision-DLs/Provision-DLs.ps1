<#
.SYNOPSIS
Create/Update Distribution Groups in Exchange Online and add members from a CSV.

.DESCRIPTION
Reads a CSV with columns: GroupName, GroupEmail, MemberEmail.
For each unique GroupEmail:
  - Ensures a Distribution Group exists (creates if missing).
  - Adds each MemberEmail to the group (skips if already a member).
Optionally creates MailContacts for external addresses that don't exist.

.PARAMETER CsvPath
Path to the CSV file. Default: .\groups.csv

.PARAMETER CreateContacts
If provided, external emails that don't exist in the tenant will be created as MailContacts.

.PARAMETER ContactOU
(Optional) The OU (on-prem synced tenants) or container where MailContacts should be created.
If omitted, Exchange Online will place them in the default container.

.EXAMPLE
.\Provision-DLs.ps1 -CsvPath .\groups.csv

.EXAMPLE
.\Provision-DLs.ps1 -CsvPath .\groups.csv -CreateContacts

.EXAMPLE
.\Provision-DLs.ps1 -CsvPath .\groups.csv -CreateContacts -ContactOU "OU=Contacts,DC=contoso,DC=com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$CsvPath = ".\groups.csv",

    [Parameter(Mandatory=$false)]
    [switch]$CreateContacts,

    [Parameter(Mandatory=$false)]
    [string]$ContactOU
)

function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing module '$Name'..." -ForegroundColor Yellow
        try {
            Install-Module $Name -Repository PSGallery -Scope CurrentUser -Force -ErrorAction Stop
        } catch {
            throw "Failed to install module '$Name': $($_.Exception.Message)"
        }
    }
    Import-Module $Name -ErrorAction Stop
}

function New-SafeAliasFromEmail {
    param([string]$Email)
    $local = ($Email -split '@')[0]
    # Allowed alias chars: letters/digits. Replace others with '-'
    ($local -replace '[^a-zA-Z0-9]', '-').Trim('-')
}

function Get-RecipientByEmail {
    param([string]$Email)
    # Get-Recipient handles SMTP addresses as identity
    try {
        return Get-Recipient -Identity $Email -ErrorAction Stop
    } catch {
        return $null
    }
}

function Ensure-Contact {
    param(
        [string]$Email,
        [string]$ContactOU
    )
    $existing = Get-RecipientByEmail -Email $Email
    if ($existing) { return $existing }

    $name = $Email
    $params = @{
        Name                  = $name
        ExternalEmailAddress  = $Email
        ErrorAction           = 'Stop'
    }
    if ($ContactOU) { $params['OrganizationalUnit'] = $ContactOU }

    Write-Host "Creating MailContact for $Email ..." -ForegroundColor Yellow
    try {
        $contact = New-MailContact @params
        Start-Sleep -Seconds 2
        return $contact
    } catch {
        Write-Warning "Failed to create MailContact for $Email $($_.Exception.Message)"
        return $null
    }
}

function Ensure-DistributionGroup {
    param(
        [string]$Name,
        [string]$PrimarySmtpAddress
    )
    # Try by SMTP first, then by display name
    $dg = $null
    try { $dg = Get-DistributionGroup -Identity $PrimarySmtpAddress -ErrorAction Stop } catch {}
    if (-not $dg) {
        try { $dg = Get-DistributionGroup -Identity $Name -ErrorAction Stop } catch {}
    }
    if ($dg) { return $dg }

    $alias = New-SafeAliasFromEmail -Email $PrimarySmtpAddress
    Write-Host "Creating Distribution Group '$Name' <$PrimarySmtpAddress> ..." -ForegroundColor Yellow
    try {
        $dg = New-DistributionGroup -Name $Name -PrimarySmtpAddress $PrimarySmtpAddress -Alias $alias -Type Distribution -ErrorAction Stop
        # Allow directory to catch up
        Start-Sleep -Seconds 3
        return $dg
    } catch {
        throw "Failed to create Distribution Group '$Name' <$PrimarySmtpAddress>: $($_.Exception.Message)"
    }
}

function Test-MemberInGroup {
    param(
        [string]$GroupIdentity,
        [string]$MemberSmtp
    )
    try {
        $members = Get-DistributionGroupMember -Identity $GroupIdentity -ResultSize Unlimited -ErrorAction Stop
        return $members | Where-Object { ($_.PrimarySmtpAddress -eq $MemberSmtp) -or ($_.WindowsEmailAddress -eq $MemberSmtp) -or ($_.Name -eq $MemberSmtp) } | ForEach-Object { $true } | Select-Object -First 1
    } catch {
        # If we can't list members (very large groups / perms), fall back to attempting add and catching the duplicate
        return $false
    }
}

# --- Prep & Connect ---
$start = Get-Date
$log = New-Object System.Collections.Generic.List[object]
$logPath = Join-Path -Path (Get-Location) -ChildPath ("DL-Provisioning-{0}.csv" -f $start.ToString("yyyyMMdd-HHmmss"))

if (-not (Test-Path $CsvPath)) { throw "CSV not found at: $CsvPath" }

# Validate headers quickly
$first = Import-Csv -Path $CsvPath | Select-Object -First 1
foreach ($h in 'GroupName','GroupEmail','MemberEmail') {
    if (-not ($first.PSObject.Properties.Name -contains $h)) {
        throw "CSV missing required column: $h"
    }
}

Ensure-Module -Name ExchangeOnlineManagement

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
} catch {
    throw "Failed to connect to Exchange Online: $($_.Exception.Message)"
}

try {
    $rows = Import-Csv -Path $CsvPath

    # Group by GroupEmail so each DG is created once
    $byGroup = $rows | Group-Object GroupEmail

    foreach ($grp in $byGroup) {
        $groupEmail = $grp.Name.Trim()
        $groupName  = ($grp.Group | Select-Object -First 1).GroupName.Trim()

        if (-not $groupEmail) { Write-Warning "Skipping a set with blank GroupEmail."; continue }
        if (-not $groupName)  { Write-Warning "GroupEmail '$groupEmail' has blank GroupName; using email local part."; $groupName = ($groupEmail -split '@')[0] }

        # Ensure the DG exists
        try {
            $dg = Ensure-DistributionGroup -Name $groupName -PrimarySmtpAddress $groupEmail
            $dgIdentity = $dg.PrimarySmtpAddress
        } catch {
            $log.Add([pscustomobject]@{
                Timestamp = (Get-Date)
                Action    = 'CreateGroup'
                Group     = $groupEmail
                Member    = ''
                Result    = 'Error'
                Detail    = $_.ToString()
            })
            Write-Error $_
            continue
        }

        # Add members
        foreach ($row in $grp.Group) {
            $memberEmail = ($row.MemberEmail).Trim()
            if (-not $memberEmail) { continue }

            $recipient = Get-RecipientByEmail -Email $memberEmail
            if (-not $recipient -and $CreateContacts) {
                $recipient = Ensure-Contact -Email $memberEmail -ContactOU $ContactOU
            }

            if (-not $recipient) {
                $msg = "Recipient not found and not created: $memberEmail"
                Write-Warning $msg
                $log.Add([pscustomobject]@{
                    Timestamp = (Get-Date)
                    Action    = 'AddMember'
                    Group     = $groupEmail
                    Member    = $memberEmail
                    Result    = 'Skipped'
                    Detail    = $msg
                })
                continue
            }

            # Skip if already a member
            if (Test-MemberInGroup -GroupIdentity $dgIdentity -MemberSmtp $memberEmail) {
                $log.Add([pscustomobject]@{
                    Timestamp = (Get-Date)
                    Action    = 'AddMember'
                    Group     = $groupEmail
                    Member    = $memberEmail
                    Result    = 'Exists'
                    Detail    = 'Already a member'
                })
                continue
            }

            try {
                Add-DistributionGroupMember -Identity $dgIdentity -Member $recipient.Identity -BypassSecurityGroupManagerCheck:$true -ErrorAction Stop
                $log.Add([pscustomobject]@{
                    Timestamp = (Get-Date)
                    Action    = 'AddMember'
                    Group     = $groupEmail
                    Member    = $memberEmail
                    Result    = 'Added'
                    Detail    = ''
                })
                Write-Host "Added $memberEmail to $groupEmail" -ForegroundColor Green
            } catch {
                # If error indicates duplicate, log as Exists
                $err = $_.Exception.Message
                if ($err -match 'is already a member' -or $err -match 'exists') {
                    $log.Add([pscustomobject]@{
                        Timestamp = (Get-Date)
                        Action    = 'AddMember'
                        Group     = $groupEmail
                        Member    = $memberEmail
                        Result    = 'Exists'
                        Detail    = $err
                    })
                } else {
                    $log.Add([pscustomobject]@{
                        Timestamp = (Get-Date)
                        Action    = 'AddMember'
                        Group     = $groupEmail
                        Member    = $memberEmail
                        Result    = 'Error'
                        Detail    = $err
                    })
                    Write-Warning "Failed to add $memberEmail to $groupEmail $err"
                }
            }
        }
    }
}
finally {
    if ($log.Count -gt 0) {
        $log | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8
        Write-Host "Summary written to: $logPath" -ForegroundColor Cyan
    }
    try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}
}

Write-Host "Done." -ForegroundColor Cyan
