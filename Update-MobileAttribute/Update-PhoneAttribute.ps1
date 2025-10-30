<#
.SYNOPSIS
Clears Mobile (mobilePhone) and sets Phone (businessPhones) from userphones.csv.

.CSV FORMAT
UPN,MobileNumber
user1@domain.com,+15551234567
user2@domain.com,+15557654321
#>

# Ensure the lightweight Graph module is present
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Write-Host "Installing Microsoft.Graph.Users module..."
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph.Users

# Connect (interactive)
Connect-MgGraph -Scopes "User.ReadWrite.All"

$csvPath = ".\userphones.csv"
if (-not (Test-Path $csvPath)) {
    Write-Error "CSV file not found at $csvPath"
    exit 1
}

$rows = Import-Csv -Path $csvPath

foreach ($row in $rows) {
    $upn   = ($row.UPN).ToString().Trim()
    $phone = ($row.MobileNumber).ToString().Trim()

    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Warning "Skipping a row with missing UPN."
        continue
    }

    try {
        if ([string]::IsNullOrWhiteSpace($phone)) {
            # No phone to set; just clear mobilePhone
            Write-Host "Clearing Mobile for $upn..."
            Update-MgUser -UserId $upn -MobilePhone "" | Out-Null
        } else
            {
            Write-Host "Clearing Mobile and setting Phone for $upn -> $phone ..."
            # Clear mobilePhone and replace businessPhones with one value
            Update-MgUser -UserId $upn -MobilePhone "" -BusinessPhones @($phone) | Out-Null
        }

        Write-Host "✅ $upn updated."
    }
    catch {
        Write-Warning "⚠️ Failed for $upn $($_.Exception.Message)"
    }
}

Disconnect-MgGraph
Write-Host "Done."
