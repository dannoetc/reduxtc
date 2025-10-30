<# 
.SYNOPSIS
  Detects and optionally removes secondary Workplace-joined accounts on Entra-joined Windows devices
  after a tenant-to-tenant migration.

.NOTES
  - Designed to run as SYSTEM via RMM.
  - Uses safe transcript handling, robust TokenBroker control, idempotent deletions, and guarded substring operations.
  - Dry-run mode supported via $RemoveSecondaryAccounts switch.
#>

Set-StrictMode -Version Latest

$FileDate = Get-Date -Format 'MM-dd-yyyy-HH-mm'

#region Config
$RemoveSecondaryAccounts = $true     # $false = discovery only (dry-run), $true = perform deletions
$Load    = $true                     # Load Settings.dat (required in production)
$UnLoad  = $true                     # Unload Settings.dat when finished
$CreateTranscripts = $true

$LogFileParentFolder = "C:\ReduxTC\Logs"
$LogFileFolder       = "C:\ReduxTC\Logs\AzureAccountCleanup"
$LogFileNameBase     = "AzureAccountCleanup"
$LogFileName         = "$LogFileNameBase$FileDate.log"
#endregion

# Ensure folders exist early
New-Item -Path $LogFileParentFolder -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
New-Item -Path $LogFileFolder       -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

# Initialize $LogPath and touch the file so first write succeeds
$script:LogPath = Join-Path $LogFileFolder $LogFileName
if (-not (Test-Path -LiteralPath $script:LogPath)) {
    New-Item -Path $script:LogPath -ItemType File -Force | Out-Null
}

#region Helpers
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERR')][string]$Level = 'INFO',
        [string]$Path = $script:LogPath
    )
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) {
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType File -Force | Out-Null
    }
    $line = "[{0}] {1}: {2}" -f (Get-Date -Format s), $Level, $Message
    Add-Content -Path $Path -Value $line -Force
}

function Stop-TransSafe {
    try { Stop-Transcript | Out-Null } catch {}
}

function New-HiddenDirectory {
    param([Parameter(Mandatory)][string]$Path)
    if (Test-Path $Path) {
        Write-Log "Path exists: $Path"
    } else {
        New-Item -Path $Path -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        try {
            (Get-Item -LiteralPath $Path).Attributes = 'Directory','Hidden'
        } catch {}
        Write-Log "Created (hidden) path: $Path"
    }
}

function Stop-AADTokenBroker {
    # Stop service and process, safely
    try {
        $svc = Get-Service -Name TokenBroker -ErrorAction Stop
        if ($svc.Status -ne 'Stopped') {
            Stop-Service -Name TokenBroker -Force -ErrorAction SilentlyContinue
            (Get-Service -Name TokenBroker -ErrorAction SilentlyContinue).WaitForStatus('Stopped','00:00:20') | Out-Null
        }
    } catch {}
    Get-Process -Name "Microsoft.AAD.BrokerPlugin" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Start-AADTokenBroker {
    try { Start-Service -Name TokenBroker -ErrorAction SilentlyContinue } catch {}
}

function Test-FileLock {
    # Returns $true if the file is locked; $false if not locked or not found
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    $oFile = New-Object System.IO.FileInfo $Path
    try {
        $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        if ($oStream) { $oStream.Close() }
        return $false
    } catch { return $true }
}

function Get-CurrentUser {
    # Best-effort: prefer console session, then fall back to CIM owner of explorer.exe
    $user = $null
    try {
        $explorers = Get-Process -Name explorer -IncludeUserName -ErrorAction SilentlyContinue
        if ($explorers) {
            $candidate = $explorers | Sort-Object { $_.SessionId -eq 1 } -Descending | Select-Object -First 1
            if ($candidate.UserName) { $user = ($candidate.UserName.Split('\') | Select-Object -Last 1) }
        }
    } catch {}
    if (-not $user) {
        $u = Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue |
             Invoke-CimMethod -MethodName GetOwner -ErrorAction SilentlyContinue |
             Select-Object -ExpandProperty User -ErrorAction SilentlyContinue
        $user = $u
    }
    return $user
}

function Get-CurrentUserSID {
    # Try CIM first
    $UserSid = $null
    try {
        $CurrentLoggedOnUser = (Get-CimInstance win32_computersystem -ErrorAction SilentlyContinue).UserName
        if ($CurrentLoggedOnUser) {
            $AdObj   = New-Object System.Security.Principal.NTAccount($CurrentLoggedOnUser)
            $strSID  = $AdObj.Translate([System.Security.Principal.SecurityIdentifier])
            $UserSid = $strSID.Value
            Write-Log "Get-CurrentUserSID: Found via CIM: $UserSid"
        }
    } catch {}

    if (-not $UserSid) {
        Write-Log "Get-CurrentUserSID: CIM null; checking explorer.exe owner" -Level WARN
        if (-not (Get-PSDrive HKU -ErrorAction SilentlyContinue)) {
            New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS | Out-Null
        }
        try {
            $owner = Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue |
                     Invoke-CimMethod -MethodName GetOwner -ErrorAction SilentlyContinue |
                     Select-Object -ExpandProperty User -ErrorAction SilentlyContinue
            $vals  = Get-ChildItem 'HKU:\*\Volatile Environment\' -ErrorAction SilentlyContinue | Get-ItemProperty -name 'USERNAME' -ErrorAction SilentlyContinue
            $ppath = $vals | Where-Object { $_.USERNAME -like "$owner" } | Select-Object -First 1 -ExpandProperty PSParentPath
            if ($ppath) {
                $UserSid = $ppath.Substring(47)  # strip 'Microsoft.PowerShell.Core\Registry::HKEY_USERS\'
                Write-Log "Get-CurrentUserSID: Found via Explorer owner: $UserSid"
            }
        } catch {}
    }
    return $UserSid
}

function Backup-SettingsDat {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$FileName
    )
    $src = Join-Path $Path 'Settings.dat'
    $dst = Join-Path $Path ("Settings-{0}.dat" -f $FileName)
    Write-Log "Backing up Settings.dat from '$src' to '$dst'"
    try {
        if (-not (Test-Path -LiteralPath $dst)) {
            Copy-Item -LiteralPath $src -Destination $dst -Force
            $s = Get-Item -LiteralPath $src -ErrorAction SilentlyContinue
            $d = Get-Item -LiteralPath $dst -ErrorAction SilentlyContinue
            Write-Log ("Backup sizes: src={0} dst={1}" -f $s.Length, $d.Length)
        } else {
            Write-Log "Backup already exists: $dst" -Level WARN
        }
    } catch {
        Write-Log "Backup failed: $($_.Exception.Message)" -Level WARN
    }
}
#endregion

#region Prep paths & transcript
New-HiddenDirectory -Path $LogFileParentFolder
New-HiddenDirectory -Path $LogFileFolder

Write-Log "Azure Account Cleanup running on $env:COMPUTERNAME"


$TranscriptPath = Join-Path $LogFileFolder "AzureAccountCleanup-Transcript-$FileDate.log"
if ($CreateTranscripts) {
    try {
        Start-Transcript -Path $TranscriptPath -ErrorAction Stop
        Write-Log "Started Transcript: $TranscriptPath"
    } catch {
        Write-Warning "Transcript could not be started: $($_.Exception.Message)"
        Write-Log "Transcript could not be started: $($_.Exception.Message)" -Level WARN
    }
}
#endregion

#region Resolve user + SID and mount HKU:
$UserSid = Get-CurrentUserSID
if (-not $UserSid) {
    Write-Warning "User SID is null; likely no interactive user. Exiting."
    Write-Log "User SID is null; exiting." -Level WARN
    Stop-TransSafe
    exit 0
}

$user = Get-CurrentUser
if (-not $user) {
    Write-Warning "Username is null; likely no interactive user. Exiting."
    Write-Log "Username is null; exiting." -Level WARN
    Stop-TransSafe
    exit 0
}

if (-not (Get-PSDrive HKU -ErrorAction SilentlyContinue)) {
    New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS | Out-Null
}
#endregion

#region Check for secondary accounts (AAD Storage value names)
$aadKeyPath = "HKU:\$UserSid\Software\Microsoft\Windows\CurrentVersion\AAD\Storage\https://login.microsoftonline.com"
if (-not (Test-Path -LiteralPath "HKU:\$UserSid")) {
    Write-Warning "HKU:\$UserSid not found, exiting."
    Write-Log "HKU:\$UserSid not found, exiting." -Level WARN
    Stop-TransSafe
    exit 1
}

$AADStorageKey = Get-ItemProperty -Path $aadKeyPath -ErrorAction SilentlyContinue
if (-not $AADStorageKey) {
    Write-Warning "No secondary accounts found in AAD Storage; exiting."
    Write-Log "No secondary accounts found; exiting."
    Stop-TransSafe
    exit 0
}

# Enumerate value names robustly (no brittle substrings)
$RegKeys   = @()
$AccountID = @()
$propNames = ($AADStorageKey | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
foreach ($name in $propNames) {
    $RegKeys   += $name
    if ($name -match '^U:(?<acct>.+)$') { $AccountID += $Matches['acct'] } else { $AccountID += $name }
}

# Find actual AAD Broker package
$PackagesRoot = "C:\Users\$user\AppData\Local\Packages"
$BrokeFolder  = Get-ChildItem -Path $PackagesRoot -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like 'Microsoft.AAD.BrokerPlugin*' } |
                Select-Object -First 1 -ExpandProperty Name

if (-not $BrokeFolder) {
    Write-Log "AAD Broker package folder not found under $PackagesRoot" -Level ERR
    Stop-TransSafe
    exit 1
}

#endregion

#region Mount Settings.dat
if ($Load) {
    $RegFileLocation = Join-Path (Join-Path $PackagesRoot $BrokeFolder) 'Settings'
    $RegFile         = Join-Path $RegFileLocation 'Settings.dat'

    Write-Log "Stopping TokenBroker & BrokerPlugin prior to mounting Settings.dat"
    Stop-AADTokenBroker

    if (Test-FileLock -Path $RegFile) {
        Write-Warning "Settings.dat locked or missing: $RegFile"
        Write-Log "Settings.dat locked or missing: $RegFile" -Level WARN
        Stop-TransSafe
        exit 1
    } else {
        Write-Log "Backing up Settings.dat"
        Backup-SettingsDat -Path $RegFileLocation -FileName $FileDate

        Write-Log "Mounting Settings.dat: $RegFile"
        & reg.exe load 'HKU\SettingsMount' "$RegFile" | Out-Null
        if (-not (Get-PSDrive HKSettingsMount -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name HKSettingsMount -PSProvider Registry -Root 'HKU\SettingsMount' | Out-Null
        }
    }
} else {
    Write-Log "Warning: Load set to false; expecting Settings.dat already mounted." -Level WARN
}
#endregion

#region Iterate accounts and compute deletions
$TokenFolder = Join-Path (Join-Path $PackagesRoot $BrokeFolder) 'AC\TokenBroker\Accounts'
$AllTokens   = @(Get-ChildItem -Path $TokenFolder -Filter *.tbacct -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
if (Test-Path -LiteralPath $TokenFolder) {
    Write-Log "Found token folder: $TokenFolder"
} else {
    Write-Log "Token folder not found: $TokenFolder" -Level WARN
}

foreach ($UniversalAccountID in $RegKeys) {

    # TenantId from SettingsMount\LocalState\SSOUsers\<U:account.tenant>
    $tenantReg = $null
    try {
        $tenantReg = & reg.exe query ("HKU\SettingsMount\LocalState\SSOUsers\{0}" -f $UniversalAccountID) /v TenantId 2>$null |
                     Where-Object { $_ -match '\s+(\S+)$' } |
                     ForEach-Object { $matches[1] -split '(?<=\G.{2})(?!$)' -replace '^','0x' -as [byte[]] }
    } catch {}

    if (-not $tenantReg) {
        Write-Log "TenantId not found for SSOUsers\$UniversalAccountID" -Level WARN
        continue
    }

    $StringTenantId = [System.Text.Encoding]::UTF8.GetString($tenantReg)
    $StringTenantId = $StringTenantId -replace '[^A-Za-z0-9:\-\.]+',''
    $StringTenantId = $StringTenantId.Substring(0, [Math]::Min(36, $StringTenantId.Length))

    # Locate UPN for this tenant (WorkplaceJoin\JoinInfo\<child>\UserEmail)
    $RegUPNObj = Get-ChildItem "HKU:\$UserSid\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\JoinInfo\*" -ErrorAction SilentlyContinue |
                 Get-ItemProperty -Name 'TenantId' -ErrorAction SilentlyContinue
    $RegUPNTempPath = $RegUPNObj | Where-Object { $_.TenantId -eq $StringTenantId } | Select-Object -First 1 -ExpandProperty PSChildName
    if (-not $RegUPNTempPath) {
        Write-Log "JoinInfo child not found for TenantId $StringTenantId" -Level WARN
        continue
    }
    $JoinInfoPath = "HKU:\$UserSid\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\JoinInfo\$RegUPNTempPath"
    $StringUPN = (Get-ItemProperty -Path $JoinInfoPath -Name "UserEmail" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty UserEmail)

    # AccountIDs to delete (map of Universal -> AccountID values)
    $AccountsIDsToDelete = @()
    $accProp = Get-ItemProperty "HKU:\SettingsMount\LocalState\AccountID" -ErrorAction SilentlyContinue
    if ($accProp) {
        $accNames = ($accProp | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
        foreach ($name in $accNames) {
            # Decode value bytes and match against Universal ID string
            $raw = & reg.exe query "HKU\SettingsMount\LocalState\AccountID" /v $name 2>$null |
                   Where-Object { $_ -match '\s+(\S+)$' } |
                   ForEach-Object { $matches[1] -split '(?<=\G.{2})(?!$)' -replace '^','0x' -as [byte[]] }
            if ($raw) {
                $TempString = [System.Text.Encoding]::UTF8.GetString($raw)
                $TempString = $TempString -replace '[^A-Za-z0-9:\-\.]+',''
                if ($TempString) {
                    if ($TempString.Length -gt 75) { $TempString = $TempString.Substring(0,75) }
                    if ($TempString -eq $UniversalAccountID) {
                        Write-Host "Matching Universal ID on AccountID: $name" -ForegroundColor Green
                        $AccountsIDsToDelete += $name
                    }
                }
            }
        }
    }

    # Tokens to delete for this UPN
    $TokensToDelete = @()
    foreach ($tok in $AllTokens) {
        try {
            $tokPath = Join-Path $TokenFolder $tok
            if (Test-Path -LiteralPath $tokPath) {
                $TokenContent = (Get-Content -LiteralPath $tokPath -Raw -ErrorAction SilentlyContinue) -replace '\u0000'
                if ($TokenContent -and $StringUPN -and ($TokenContent -like "*$StringUPN*")) {
                    Write-Host "Found matching Token for $StringUPN : $tok" -ForegroundColor Green
                    $TokensToDelete += $tok
                }
            }
        } catch {}
    }

    # Log what we found
    Write-Host  "On Account $UniversalAccountID"
    Write-Log   "On Account $UniversalAccountID"
    Write-Host  "Account UPN: $StringUPN"
    Write-Log   "Account UPN: $StringUPN"
    Write-Host  "Tenant ID : $StringTenantId"
    Write-Log   "Tenant ID : $StringTenantId"

    $ssoPath   = "HKU:\SettingsMount\LocalState\SSOUsers\$UniversalAccountID"
    $utaidPath = "HKU:\SettingsMount\LocalState\UniversalToAccountID"
    Write-Log  ("Found SSO path: {0}" -f (Test-Path -LiteralPath $ssoPath))

    if (Get-ItemProperty -Path $utaidPath -Name $UniversalAccountID -ErrorAction SilentlyContinue) {
        Write-Log "Found UniversalToAccountID value: $utaidPath\$UniversalAccountID"
    } else {
        Write-Log "UniversalToAccountID value not found: $utaidPath\$UniversalAccountID" -Level WARN
    }

    $aadStorePath = $aadKeyPath
    if (Get-ItemProperty -Path $aadStorePath -Name $UniversalAccountID -ErrorAction SilentlyContinue) {
        Write-Log "Found AAD Storage value: $aadStorePath\$UniversalAccountID"
    } else {
        Write-Log "AAD Storage value not found: $aadStorePath\$UniversalAccountID" -Level WARN
    }

    $acctPicPath = "HKU:\SettingsMount\LocalState\AccountPicture"
    if (Get-ItemProperty -Path $acctPicPath -Name $StringUPN -ErrorAction SilentlyContinue) {
        Write-Log "Found AccountPicture: $acctPicPath\$StringUPN"
    } else {
        Write-Log "AccountPicture not found: $acctPicPath\$StringUPN" -Level WARN
    }
    if (Get-ItemProperty -Path $acctPicPath -Name "$($StringUPN)|perUser" -ErrorAction SilentlyContinue) {
        Write-Log "Found AccountPicture (perUser): $acctPicPath\$StringUPN|perUser"
    } else {
        Write-Log "AccountPicture (perUser) not found: $acctPicPath\$StringUPN|perUser" -Level WARN
    }

    $tenantInfoPath = "HKU:\$UserSid\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\TenantInfo\$StringTenantId"
    if (Test-Path -LiteralPath $tenantInfoPath) {
        Write-Log "Found TenantInfo key: $tenantInfoPath"
    } else {
        Write-Log "TenantInfo key not found: $tenantInfoPath" -Level WARN
    }

    Write-Log "Also to be deleted: HKU:\$UserSid\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\JoinInfo\$RegUPNTempPath"

    foreach ($X in $AccountsIDsToDelete) {
        if ($X) {
            if (Get-ItemProperty -Path "HKU:\SettingsMount\LocalState\AccountID" -Name $X -ErrorAction SilentlyContinue) {
                Write-Log "AccountID to remove: HKU:\SettingsMount\LocalState\AccountID\$X"
            } else {
                Write-Log "AccountID not found: HKU:\SettingsMount\LocalState\AccountID\$X" -Level WARN
            }
        }
    }

    if (Test-Path -LiteralPath $TokenFolder) {
        Write-Log "Token folder present: $TokenFolder"
    } else {
        Write-Log "Token folder not present: $TokenFolder" -Level WARN
    }

    foreach ($Z in $TokensToDelete) {
        $tokPath = Join-Path $TokenFolder $Z
        if (Test-Path -LiteralPath $tokPath) {
            Write-Log "Token to remove: $tokPath"
        } else {
            Write-Log "Token not found: $Z" -Level WARN
        }
    }

    # Deletions
    $DoDelete = [bool]$RemoveSecondaryAccounts
    if ($DoDelete) {
        Write-Log "Deletion is ENABLED, starting."

        Remove-Item -Path $ssoPath -Recurse -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $utaidPath -Name $UniversalAccountID -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $acctPicPath -Name $StringUPN -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $acctPicPath -Name "$($StringUPN)|perUser" -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $aadStorePath -Name $UniversalAccountID -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $tenantInfoPath -Recurse -Force -ErrorAction SilentlyContinue

        $joinInfoDel = "HKU:\$UserSid\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\JoinInfo\$RegUPNTempPath"
        Remove-Item -Path $joinInfoDel -Recurse -Force -ErrorAction SilentlyContinue

        foreach ($X in $AccountsIDsToDelete) {
            if ($X) {
                Remove-ItemProperty -Path "HKU:\SettingsMount\LocalState\AccountID" -Name $X -Force -ErrorAction SilentlyContinue
            }
        }

        foreach ($Z in $TokensToDelete) {
            $tokPath = Join-Path $TokenFolder $Z
            if (Test-Path -LiteralPath $tokPath) {
                Remove-Item -Path $tokPath -Force -ErrorAction SilentlyContinue
            } else {
                Write-Log "Not Found Token: $Z" -Level WARN
            }
        }

        Write-Log "Deletion complete."
    } else {
        Write-Log "Deletion is DISABLED (dry-run)."
    }
}
#endregion

#region Unmount Settings.dat
if ($Unload) {
    Write-Log "Unloading Settings.dat"
    if (Get-PSDrive -Name HKSettingsMount -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name HKSettingsMount -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    Start-Sleep -Seconds 10
    & reg.exe unload 'HKU\SettingsMount' | Out-Null
} else {
    Write-Log "Unload of Settings.dat is DISABLED" -Level WARN
}
#endregion

#region Cleanup & exit
Start-AADTokenBroker
Write-Log "Script ending."
Stop-TransSafe
exit 0
#endregion
