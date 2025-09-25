<# 
.SYNOPSIS
  Pre-provision (enqueue) OneDrive personal sites for users listed in a CSV.

.DESCRIPTION
  Reads .\useronedrive.csv (or a provided path), extracts user emails/UPNs,
  and calls Request-SPOPersonalSite in batches. Generates a results log.

.PARAMETER AdminUrl
  Your SPO Admin URL, e.g. https://contoso-admin.sharepoint.com

.PARAMETER CsvPath
  Path to the input CSV. Default: .\useronedrive.csv

.PARAMETER BatchSize
  Number of users per Request-SPOPersonalSite call (max recommended ~200). Default: 100

.PARAMETER NoWait
  If set, adds -NoWait so calls return immediately after enqueueing.

.NOTES
  Requires Microsoft.Online.SharePoint.PowerShell
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$AdminUrl,  # e.g. https://contoso-admin.sharepoint.com

    [string]$CsvPath = ".\useronedrive.csv",

    [ValidateRange(1,200)]
    [int]$BatchSize = 100,

    [switch]$NoWait
)

function Ensure-Module {
    param(
        [Parameter(Mandatory=$true)][string]$Name
    )
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Module '$Name' not found. Installing from PSGallery..." -ForegroundColor Yellow
        try {
            $null = Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            throw "Failed to install module '$Name': $($_.Exception.Message)"
        }
    }
    Import-Module $Name -ErrorAction Stop
}

function Get-EmailsFromCsv {
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        throw "CSV not found at '$Path'."
    }

    $data = Import-Csv -Path $Path
    if (-not $data -or $data.Count -eq 0) {
        throw "CSV '$Path' appears to be empty."
    }

    # Accept common header names
    $candidateHeaders = @('UPN','UserPrincipalName','UserEmail','Email')
    $header = $candidateHeaders | Where-Object { $_ -in $data[0].PSObject.Properties.Name }

    if (-not $header) {
        throw "CSV must have one of these headers: $($candidateHeaders -join ', '). Found: $($data[0].PSObject.Properties.Name -join ', ')"
    }

    # Basic email/UPN validation
    $emailRegex = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

    $emails =
        $data |
        ForEach-Object { $_.$header } |
        Where-Object { $_ -and ($_ -match $emailRegex) } |
        ForEach-Object { $_.Trim().ToLower() } |
        Sort-Object -Unique

    if (-not $emails -or $emails.Count -eq 0) {
        throw "No valid emails/UPNs found in CSV '$Path'."
    }

    return $emails
}

function Invoke-OneDriveProvisioning {
    param(
        [Parameter(Mandatory=$true)][string[]]$Emails,
        [Parameter(Mandatory=$true)][string]$AdminUrl,
        [int]$BatchSize = 100,
        [switch]$NoWait
    )

    $results = New-Object System.Collections.Generic.List[object]

    # Connect to SPO Admin
    Write-Host "Connecting to $AdminUrl ..." -ForegroundColor Cyan
    try {
        Connect-SPOService -Url $AdminUrl -ErrorAction Stop
    } catch {
        throw "Failed to connect to SharePoint Online Admin: $($_.Exception.Message)"
    }

    # Batch the requests
    $total = $Emails.Count
    $batches = [Math]::Ceiling($total / $BatchSize)
    Write-Host "Queueing OneDrive provisioning for $total user(s) in $batches batch(es)..." -ForegroundColor Cyan

    for ($i = 0; $i -lt $total; $i += $BatchSize) {
        $batch = $Emails[$i..([Math]::Min($i + $BatchSize - 1, $total - 1))]

        $idx = [int]([Math]::Floor($i / $BatchSize)) + 1
        Write-Host ("Batch {0}/{1}: {2} user(s)" -f $idx, $batches, $batch.Count) -ForegroundColor Green

        try {
            if ($NoWait) {
                Request-SPOPersonalSite -UserEmails $batch -NoWait -ErrorAction Stop
            } else {
                Request-SPOPersonalSite -UserEmails $batch -ErrorAction Stop
            }

            foreach ($u in $batch) {
                $results.Add([pscustomobject]@{
                    Timestamp = (Get-Date).ToString("s")
                    UserEmail = $u
                    Status    = "Enqueued"
                    Details   = if ($NoWait) { "NoWait" } else { "Queued" }
                })
            }
        } catch {
            $err = $_.Exception.Message
            Write-Warning "Batch $idx failed: $err"

            foreach ($u in $batch) {
                $results.Add([pscustomobject]@{
                    Timestamp = (Get-Date).ToString("s")
                    UserEmail = $u
                    Status    = "Error"
                    Details   = $err
                })
            }
        }
    }

    return $results
}

# --- Main ---
try {
    Ensure-Module -Name "Microsoft.Online.SharePoint.PowerShell"

    $emails = Get-EmailsFromCsv -Path $CsvPath

    $results = Invoke-OneDriveProvisioning -Emails $emails -AdminUrl $AdminUrl -BatchSize $BatchSize -NoWait:$NoWait

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $outFile = Join-Path -Path (Get-Location) -ChildPath "OneDriveProvisioningResults-$stamp.csv"
    $results | Export-Csv -NoTypeInformation -Path $outFile

    Write-Host ""
    Write-Host "Done. Results saved to: $outFile" -ForegroundColor Cyan
    Write-Host "Note: Provisioning is enqueued. Actual site creation can take time after the request." -ForegroundColor Yellow
}
catch {
    Write-Error $_
    exit 1
}
