# Detect SecureBoot
$ErrorActionPreference = "Ignore"
Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
if (-not(Get-PackageProvider -name nuget)) {
	Install-PackageProvider -Name NuGet -Confirm:$false
	Start-Sleep -Seconds 60
	}

if (-not(Get-InstalledModule -Name DellBIOSProvider)) {
	Install-Module -Name DellBIOSProvider -Confirm:$false
	Start-Sleep -Seconds 60
	}

Import-Module DellBIOSProvider -force

$SecureBoot = Get-Item -Path DellSmbios:\SecureBoot\SecureBoot | select currentvalue

if ($SecureBoot -eq "Enabled") { 
	exit 0
	}
	
else { 
	exit 1
	}