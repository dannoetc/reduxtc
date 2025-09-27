<# 
    Enable BitLocker on OS drive and escrow recovery key to C:\ReduxTC\Bitlocker\BitLockerKey.txt
    Also writes a detailed log to C:\ReduxTC\Bitlocker\BitLocker.log
    Intended to run as SYSTEM on Intune/Entra-joined Windows 10/11 (Pro/Enterprise).
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -------- Paths --------
$baseFolder   = 'C:\ReduxTC'
$bitlockerDir = Join-Path $baseFolder 'Bitlocker'
$escrowFile   = Join-Path $bitlockerDir 'BitLockerKey.txt'
$logFile      = Join-Path $bitlockerDir 'BitLocker.log'

# -------- Helpers --------
function Ensure-Folder { param([string]$Path) if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null } }
function Secure-ItemAcl { param([string]$Path) & icacls $Path /inheritance:r | Out-Null; & icacls $Path /grant:r "SYSTEM:(F)" "Administrators:(F)" | Out-Null }
function To-Array { param($Value) if ($null -eq $Value) { @() } elseif ($Value -is [System.Array]) { $Value } else { @($Value) } }
function Count-Of { param($x) @($x).Count }

function Write-Log {
    param([string]$Message,[string]$Level = "INFO")
    $stamp = Get-Date -Format s
    $line  = "[${stamp}] [$Level] $Message"
    Write-Output $line
    try {
        if (-not (Test-Path -LiteralPath $logFile)) { $null = New-Item -ItemType File -Path $logFile -Force; Secure-ItemAcl -Path $logFile }
        Add-Content -LiteralPath $logFile -Value $line
    } catch {}
}

function Get-OsVolume {
    $vol = Get-BitLockerVolume | Where-Object { $_.VolumeType -eq 'OperatingSystem' }
    if (-not $vol) { throw "Could not locate the operating system BitLocker volume." }
    return $vol
}

function Get-RecoveryPasswordsFromManageBde {
    param([string]$MountPoint)
    $result = @()
    $out = & manage-bde -protectors -get $MountPoint 2>$null
    if (-not $out) { return $result }

    for ($i=0; $i -lt $out.Count; $i++) {
        if ($out[$i] -match 'Numerical Password') {
            $id = $null; $pwd = $null
            for ($j=$i; $j -lt $out.Count; $j++) {
                if (-not $id -and $out[$j] -match 'ID:\s+\{([0-9A-Fa-f\-]+)\}') { $id = $Matches[1] }
                if ($out[$j] -match 'Password:\s*$') { if ($j + 1 -lt $out.Count) { $pwd = $out[$j+1].Trim() }; break }
                if ($out[$j].Trim() -eq '' -and $j -gt $i) { break }
            }
            if ($id -and $pwd) {
                $result += [pscustomobject]@{ KeyProtectorId = $id; RecoveryPassword = $pwd }
            }
        }
    }
    return $result
}

function Get-RecoveryProtectorIds {
    param($KeyProtectors)
    (To-Array $KeyProtectors) | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | ForEach-Object { $_.KeyProtectorId }
}

function Try-BackupToAAD {
    param([string]$MountPoint,[string[]]$KeyProtectorIds)
    try {
        if (Get-Command BackupToAAD-BitLockerKeyProtector -ErrorAction SilentlyContinue) {
            foreach ($kid in To-Array $KeyProtectorIds) {
                BackupToAAD-BitLockerKeyProtector -MountPoint $MountPoint -KeyProtectorId $kid -ErrorAction Stop | Out-Null
                Write-Log "Backed up protector {$kid} to Azure AD."
            }
        } else { Write-Log "BackupToAAD-BitLockerKeyProtector not available on this system." }
    } catch { Write-Log "Azure AD backup attempt failed: $($_.Exception.Message)" "WARN" }
}

# -------- Main --------
try {
    Ensure-Folder $baseFolder
    Ensure-Folder $bitlockerDir
    Secure-ItemAcl -Path $bitlockerDir

    Write-Log "Starting BitLocker enable & escrow."

    if (-not (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue)) { throw "BitLocker cmdlets not available. Windows edition may not support BitLocker." }

    $os    = Get-OsVolume
    $mount = $os.MountPoint

    # TPM status
    $tpmPresent = $false; $tpmReady = $false
    if (Get-Command Get-Tpm -ErrorAction SilentlyContinue) { $tpm = Get-Tpm; $tpmPresent = [bool]$tpm.TpmPresent; $tpmReady = [bool]$tpm.TpmReady }
    Write-Log "TPM present: $tpmPresent, TPM ready: $tpmReady"

    $didEnable = $false
    $newProtectorIds = @()
    $capturedRecovery = @()

    # Common enable params
    $common = @{
        MountPoint       = $mount
        EncryptionMethod = 'XtsAes256'
        UsedSpaceOnly    = $true
        SkipHardwareTest = $true
    }

    if ($os.ProtectionStatus -eq 'On') {
        Write-Log "BitLocker already ON for $mount. Ensuring a Recovery Password exists."
        $current = Get-OsVolume
        $beforeIds = Get-RecoveryProtectorIds $current.KeyProtector

        if ((Count-Of $beforeIds) -eq 0) {
            # Add recovery protector; suppress noisy warning output
            $null = Add-BitLockerKeyProtector -MountPoint $mount -RecoveryPasswordProtector -ErrorAction Stop -WarningAction SilentlyContinue -Confirm:$false

            $afterIds = Get-RecoveryProtectorIds (Get-OsVolume).KeyProtector
            $newIds   = (To-Array $afterIds) | Where-Object { (To-Array $beforeIds) -notcontains $_ }
            $newProtectorIds += $newIds
        }
    } else {
        Write-Log "BitLocker is OFF on $mount. Enabling…"
        if ($tpmPresent -and $tpmReady) {
            Write-Log "Using TPM protector, then adding Recovery Password."
            Enable-BitLocker @common -TpmProtector

            $beforeIds = Get-RecoveryProtectorIds (Get-OsVolume).KeyProtector
            $null = Add-BitLockerKeyProtector -MountPoint $mount -RecoveryPasswordProtector -ErrorAction Stop -WarningAction SilentlyContinue -Confirm:$false
            $afterIds = Get-RecoveryProtectorIds (Get-OsVolume).KeyProtector
            $newIds   = (To-Array $afterIds) | Where-Object { (To-Array $beforeIds) -notcontains $_ }
            $newProtectorIds += $newIds
        } else {
            Write-Log "TPM not present/ready. Enabling with Recovery Password protector."
            # This prints a warning with the password; silence it.
            $null = Enable-BitLocker @common -RecoveryPasswordProtector -WarningAction SilentlyContinue
        }

        Start-Sleep -Seconds 3
        $post = Get-OsVolume
        if ($post.ProtectionStatus -eq 'On' -or $post.ProtectionStatus -eq 1) {
            $didEnable = $true
            Write-Log "Enable initiated. Encryption: $($post.EncryptionPercentage)% (background)."
        } else {
            Write-Log "Enable initiated; status not yet 'On' (expected to flip soon)." "WARN"
        }
    }

    # Build escrow entries:
    # Parse *all* recovery passwords then (if we created new protectors) label those as new
    $parsed = To-Array (Get-RecoveryPasswordsFromManageBde -MountPoint $mount)

    # Prefer including all recovery passwords (so you have the complete set); de-dup by ID
    $seen = @{}
    $recoveryEntries = foreach ($p in $parsed) {
        if (-not $seen.ContainsKey($p.KeyProtectorId)) { $seen[$p.KeyProtectorId] = $true; $p }
    }

    # Compose escrow text
    $deviceName = $env:COMPUTERNAME
    $ts      = Get-Date -Format "yyyy-MM-dd HH:mm:ss zzz"
    $tpmNote = if ($tpmPresent -and $tpmReady) { "TPM: Present & Ready" } elseif ($tpmPresent) { "TPM: Present, NOT Ready" } else { "TPM: Not Present" }

    $lines = @()
    $lines += "==== BitLocker Escrow ===="
    $lines += "Device:        $deviceName"
    $lines += "OS Volume:     $mount"
    $lines += "Timestamp:     $ts"
    $lines += "TPM Status:    $tpmNote"
    $lines += "Protection:    $((Get-OsVolume).ProtectionStatus)"
    if ($didEnable) { $lines += "Action:        Enabled BitLocker (UsedSpaceOnly, XtsAes256, SkipHardwareTest)" }
    $lines += ""

    if ((Count-Of $recoveryEntries) -gt 0) {
        $lines += "Recovery Password(s):"
foreach ($entry in (To-Array $recoveryEntries)) {
    $newTag = ""
        if ((To-Array $newProtectorIds) -contains $entry.KeyProtectorId) {
        $newTag = " [NEW]"
        }
        $lines += "  ID: {$($entry.KeyProtectorId)}$newTag"
        $lines += "  Password: $($entry.RecoveryPassword)"
        $lines += ""
    }
    } else {
        $lines += "No Recovery Passwords found."
        $lines += ""
    }

    # Write/append escrow file
    if (Test-Path -LiteralPath $escrowFile) {
        Add-Content -LiteralPath $escrowFile -Value ($lines -join [Environment]::NewLine)
        Add-Content -LiteralPath $escrowFile -Value ("-"*60)
    } else {
        Set-Content -LiteralPath $escrowFile -Value ($lines -join [Environment]::NewLine) -Force
        Secure-ItemAcl -Path $escrowFile
    }
    Write-Log "Escrow written: $escrowFile"

    # AAD backup for newly created protectors (if any)
    $newProtectorIds = To-Array $newProtectorIds
    if ((Count-Of $newProtectorIds) -gt 0) { 
        $BLV = Get-BitLockerVolume -MountPoint "C:"
        BackupToAAD-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $BLV.KeyProtector[1].KeyProtectorId 
    }
    
    Import-Module $env:SyncroModule 
    Log-Activity -Message "BitLocker enabled, key escrowed." -EventName "BitLocker Enabled"
    
    Upload-File -Filepath $escrowFile 
    Write-Log "Completed successfully."
    Exit 0

} catch {
    $msg = "ERROR: $($_.Exception.Message)"
    Write-Error $msg
    Write-Log $msg "ERROR"
    try {
        Ensure-Folder $bitlockerDir; Secure-ItemAcl -Path $bitlockerDir
        if (-not (Test-Path -LiteralPath $logFile)) { $null = New-Item -ItemType File -Path $logFile -Force; Secure-ItemAcl -Path $logFile }
        Add-Content -LiteralPath $logFile -Value ("[{0}] [ERROR] {1}" -f (Get-Date -Format s), $msg)
    } catch {}
    Exit 1
}
