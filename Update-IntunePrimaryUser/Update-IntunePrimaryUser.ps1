# Connect to Microsoft Graph
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All", "User.Read.All"

function Get-LastLoggedOnUser {
    param (
        [string]$DeviceId
    )
    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$DeviceId')?`$select=id,deviceName,usersLoggedOn"
        $device = Invoke-MgGraphRequest -Uri $uri -Method GET

        if ($device.usersLoggedOn -and $device.usersLoggedOn.Count -gt 0) {
            # Sort users by last login time and get the most recent one
            $lastUser = $device.usersLoggedOn | Sort-Object -Property lastLogOnDateTime -Descending | Select-Object -First 1
            return $lastUser.userId
        }
        return $null
    }
    catch {
        Write-Host "Error getting last logged on user for device $DeviceId : $_" -ForegroundColor Red
        return $null
    }
}

function Set-PrimaryUser {
    param (
        [string]$DeviceId,
        [string]$UserId
    )
    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$DeviceId')/users/`$ref"
        $body = @{
            "@odata.id" = "https://graph.microsoft.com/beta/users('$UserId')"
        } | ConvertTo-Json

        Invoke-MgGraphRequest -Uri $uri -Method POST -Body $body -ContentType "application/json"
        return $true
    }
    catch {
        Write-Host "Error setting primary user for device $DeviceId : $_" -ForegroundColor Red
        return $false
    }
}

# Get all Windows Intune Managed Devices
$devices = Get-MgDeviceManagementManagedDevice -Filter "operatingSystem eq 'Windows'"

foreach ($device in $devices) {
    Write-Host "Processing device: $($device.DeviceName)" -ForegroundColor Cyan

    # Get the last logged on user
    $lastUserId = Get-LastLoggedOnUser -DeviceId $device.Id
    if ($lastUserId) {
        $lastUser = Get-MgUser -UserId $lastUserId
        Write-Host "Last logged on user: $($lastUser.UserPrincipalName)" -ForegroundColor Yellow

        # Check if the current primary user is different from the last logged on user
        if ($device.UserPrincipalName -ne $lastUser.UserPrincipalName) {
            Write-Host "Updating primary user..." -ForegroundColor Yellow
            $result = Set-PrimaryUser -DeviceId $device.Id -UserId $lastUserId
            if ($result) {
                Write-Host "Primary user updated successfully to $($lastUser.UserPrincipalName)" -ForegroundColor Green
            }
            else {
                Write-Host "Failed to update primary user" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Primary user is already set to the last logged on user" -ForegroundColor Green
        }
    }
    else {
        Write-Host "No last logged on user found for this device" -ForegroundColor Yellow
    }

    Write-Host ""
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph