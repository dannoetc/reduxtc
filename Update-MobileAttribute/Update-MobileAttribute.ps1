<#
.SYNOPSIS
Updates the Mobile phone attribute for users in Azure AD from a CSV file.

.DESCRIPTION
This script imports a CSV file named "userphones.csv" containing UPN and MobileNumber fields.
It connects to Microsoft Graph with the necessary permissions and updates each user's Mobile attribute.

.EXAMPLE
.\Update-MobileAttribute.ps1
#>

# Ensure Microsoft Graph Users module is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Write-Host "Installing Microsoft.Graph.Users module..."
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Users

# Connect to Microsoft Graph interactively
Connect-MgGraph -Scopes "User.ReadWrite.All"

# Import CSV
$csvPath = ".\userphones.csv"

if (-not (Test-Path $csvPath)) {
    Write-Error "CSV file not found at $csvPath"
    exit
}

$users = Import-Csv -Path $csvPath

foreach ($user in $users) {
    $upn = $user.UPN
    $mobile = $user.MobileNumber

    if ([string]::IsNullOrWhiteSpace($upn) -or [string]::IsNullOrWhiteSpace($mobile)) {
        Write-Warning "Skipping entry with missing UPN or MobileNumber."
        continue
    }

    try {
        Write-Host "Updating mobile number for $upn to $mobile..."
        Update-MgUser -UserId $upn -MobilePhone $mobile
        Write-Host "✅ Successfully updated $upn"
    }
    catch {
        Write-Warning "⚠️ Failed to update $upn $($_.Exception.Message)"
    }
}

Write-Host "All updates processed."
Disconnect-MgGraph
