<#
Pull BitLocker Recovery Key from Microsoft Graph (Intune/Entra) and
write it to the Syncro custom field "BitLockerKey" using the doc flow:
  1) List recoveryKeys
  2) Pick the row for this deviceId (newest)
  3) GET .../recoveryKeys/{id}?$select=key

RUN AS: SYSTEM (RMM/Intune), 64-bit PowerShell 5.1+
REQUIRES: App registration with Application permission BitLockerKey.Read.All (admin-consented).
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Ensure modern TLS
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.ServicePointManager]::SecurityProtocol } catch {}

# ====== >>>>>>>  SUPPLY YOUR CREDENTIALS  <<<<<<< ======
$TenantId     = "ef0d09ea-dcbd-4116-8af0-6e16a5351c7a"
$ClientId     = "5935284d-44f2-4def-8c5e-ec2ee418dcdf"
$ClientSecret = "Mu28Q~~coWwy34IvNr2lIuzCOjIY5s0lrm9o1dek"

# ====== Paths & Logging ======
$BaseFolder   = 'C:\ReduxTC'
$OutDir       = Join-Path $BaseFolder 'Bitlocker'
$LogPath      = Join-Path $OutDir 'GraphKeyPull.log'

function Ensure-Folder { param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}
function Secure-ItemAcl { param([string]$Path)
    & icacls $Path /inheritance:r | Out-Null
    & icacls $Path /grant:r "SYSTEM:(F)" "Administrators:(F)" | Out-Null
}
function Write-Log { param([string]$Message, [string]$Level="INFO")
    $stamp = Get-Date -Format s
    $line  = "[${stamp}] [$Level] $Message"
    Write-Output $line
    try {
        if (-not (Test-Path -LiteralPath $LogPath)) { New-Item -ItemType File -Path $LogPath -Force | Out-Null; Secure-ItemAcl $LogPath }
        Add-Content -LiteralPath $LogPath -Value $line
    } catch {}
}

# ====== Utilities ======
function Normalize-DeviceId {
    param([string]$s)
    if (-not $s) { return $null }
    $t = $s.Trim('{}').Trim()
    $out = [Guid]::Empty
    if ([Guid]::TryParse($t, [ref]$out)) { return $out.ToString() }
    if ($t -match '^[0-9a-fA-F]{32}$') {
        $g = ($t.Substring(0,8) + '-' + $t.Substring(8,4) + '-' + $t.Substring(12,4) + '-' + $t.Substring(16,4) + '-' + $t.Substring(20,12))
        if ([Guid]::TryParse($g, [ref]$out)) { return $out.ToString() }
    }
    return $null
}

function Get-LocalAadDeviceId {
    # 1) dsregcmd /status
    try {
        $out = & dsregcmd /status 2>$null
        if ($out) {
            $text = ($out | Out-String)
            if ($text -match 'DeviceId\s*:\s*([0-9a-fA-F\-]{36})') {
                $norm = Normalize-DeviceId $matches[1]
                if ($norm) { return $norm }
            }
        }
    } catch {}
    # 2) CloudDomainJoin\JoinInfo
    try {
        $joinKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo'
        if (Test-Path $joinKey) {
            foreach ($k in (Get-ChildItem -LiteralPath $joinKey -ErrorAction Stop)) {
                $norm = Normalize-DeviceId $k.PSChildName
                if ($norm) { return $norm }
            }
        }
    } catch {}
    # 3) Enrollment values
    try {
        $enrollRoot = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
        if (Test-Path $enrollRoot) {
            foreach ($sub in (Get-ChildItem $enrollRoot -ErrorAction SilentlyContinue)) {
                foreach ($name in 'AADDeviceID','AzureAdDeviceId','AadDeviceId') {
                    try {
                        $val = (Get-ItemProperty -LiteralPath $sub.PSPath -Name $name -ErrorAction SilentlyContinue).$name
                        $norm = Normalize-DeviceId $val
                        if ($norm) { return $norm }
                    } catch {}
                }
            }
        }
    } catch {}
    return $null
}

function Get-GraphToken {
    param(
        [Parameter(Mandatory)] [string] $TenantId,
        [Parameter(Mandatory)] [string] $ClientId,
        [Parameter(Mandatory)] [string] $ClientSecret
    )
    $uri  = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }
    Invoke-RestMethod -Method POST -Uri $uri -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
}

function Invoke-GraphJson {
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST','PATCH','PUT','DELETE')] [string]$Method,
        [Parameter(Mandatory)] [string]$Uri,
        [Parameter(Mandatory)] [hashtable]$Headers
    )
    try {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -ErrorAction Stop
    } catch {
        # PS 5.1-safe extraction of Graph error text
        $detail = $null
        if ($null -ne $_.ErrorDetails) {
            try { if ($_.ErrorDetails.Message) { $detail = $_.ErrorDetails.Message } } catch {}
        }
        if (-not $detail -and $null -ne $_.Exception -and $null -ne $_.Exception.Response) {
            try {
                $respStream = $_.Exception.Response.GetResponseStream()
                if ($respStream) {
                    $sr = New-Object System.IO.StreamReader($respStream)
                    $detail = $sr.ReadToEnd()
                    $sr.Dispose(); $respStream.Dispose()
                }
            } catch {}
        }
        if ($detail -and $detail.Trim().StartsWith('{')) {
            try {
                $j = $detail | ConvertFrom-Json -ErrorAction Stop
                if ($j.error -and $j.error.message) { $detail = $j.error.message }
            } catch {}
        }
        if (-not $detail) { $detail = $_.Exception.Message }
        throw "Graph call failed: $Method $Uri`n$detail"
    }
}

function New-GraphHeaders {
    param([Parameter(Mandatory)][string]$AccessToken)
    return @{
        Authorization        = "Bearer $AccessToken"
        'User-Agent'         = 'Redux-BitLockerKeyPull/1.0'  # Required by this API
        'Accept'             = 'application/json'
        'ConsistencyLevel'   = 'eventual'
        'Prefer'             = 'odata.maxpagesize=999'
        'ocp-client-name'    = 'Redux-GraphKeyPull'
        'ocp-client-version' = '1.0'
    }
}

function Get-DeviceRecoveryKeyFromGraph {
    param(
        [Parameter(Mandatory)] [string] $AccessToken,
        [Parameter(Mandatory)] [string] $DeviceIdGuid
    )

    $base    = "https://graph.microsoft.com/v1.0/informationProtection/bitlocker"
    $headers = New-GraphHeaders -AccessToken $AccessToken

    $allForDevice = @()
    $next = "$base/recoveryKeys"

    do {
        $page = Invoke-GraphJson -Method GET -Uri $next -Headers $headers

        # ------- Normalize page to an items array -------
        $items = @()
        if ($null -ne $page) {
            if ($page -is [System.Array]) {
                # Top-level JSON array
                $items = @($page)
            } elseif ($page.PSObject -and $page.PSObject.Properties['value']) {
                # Standard Graph collection { value: [...] }
                $items = @($page.value)
            } else {
                # Singleton object fallback
                $items = @($page)
            }
        }

        # ------- Collect items for this device -------
        if ($items.Count -gt 0) {
            $matches = $items | Where-Object {
                $_.deviceId -and ([string]$_.deviceId).ToLower() -eq $DeviceIdGuid.ToLower()
            }
            if ($matches) { $allForDevice += $matches }
        }

        # ------- Advance paging only if '@odata.nextLink' actually exists -------
        $next = $null
        if ($page -and $page.PSObject -and $page.PSObject.Properties['@odata.nextLink']) {
            $next = [string]$page.'@odata.nextLink'
        }

    } while ($next)

    if ($allForDevice.Count -eq 0) { return $null }

    # Prefer newest key (createdDateTime can be string; cast to [datetime] for reliable ordering)
    $chosen = $allForDevice |
        Sort-Object @{ Expression = { [datetime]$_.createdDateTime }; Descending = $true } |
        Select-Object -First 1

    if (-not $chosen -or -not $chosen.id) { return $null }

    # Fetch the actual key (audited) using $select=key
    $keyUri = ('{0}/recoveryKeys/{1}?%24select=key' -f $base, $chosen.id)

    $keyObj = Invoke-GraphJson -Method GET -Uri $keyUri -Headers $headers
    if ($keyObj.key) {
        return [pscustomobject]@{
            RecoveryKey = $keyObj.key
            RecoveryId  = $chosen.id
            VolumeType  = $chosen.volumeType
            Created     = $chosen.createdDateTime
        }
    } else {
        return $null
    }
}


function Update-SyncroBitLockerField {
    param([string]$RecoveryKey)
    try {
        Import-Module $env:SyncroModule -ErrorAction Stop
        if (Get-Command Set-Asset-Field -ErrorAction SilentlyContinue) {
            Set-Asset-Field -Name "BitLockerKey" -Value $RecoveryKey
            Write-Log "Updated Syncro custom field 'BitLockerKey'."
            return $true
        } else {
            Write-Log "Syncro module loaded but Set-Asset-Field not found." "WARN"
            return $false
        }
    } catch {
        Write-Log "Could not update Syncro field: $($_.Exception.Message)" "WARN"
        return $false
    }
}

# ====== Main ======
try {
    Ensure-Folder $BaseFolder
    Ensure-Folder $OutDir
    Secure-ItemAcl $OutDir

    Write-Log "Starting Graph BitLocker key retrieval."

    if ($TenantId -match '^<YOUR_' -or $ClientId -match '^<YOUR_' -or $ClientSecret -match '^<YOUR_') {
        throw "Please set TenantId/ClientId/ClientSecret at the top of the script."
    }

    $deviceId = Get-LocalAadDeviceId
    if (-not $deviceId) { throw "Azure AD deviceId not found or not a valid GUID on this machine; is it Entra-joined?" }
    Write-Log "Local Azure AD deviceId (GUID): $deviceId"

    $tokenResp = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    if (-not $tokenResp.access_token) { throw "Failed to obtain Graph token." }
    $accessToken = $tokenResp.access_token
    Write-Log "Graph token acquired."

    $rk = Get-DeviceRecoveryKeyFromGraph -AccessToken $accessToken -DeviceIdGuid $deviceId
    if (-not $rk) {
        Write-Log "No BitLocker recovery key found in Graph for this device." "WARN"
        # Optional: Update-SyncroBitLockerField -RecoveryKey ""
        exit 0
    }

    # Mask most of the key in logs
    $masked = $rk.RecoveryKey -replace '^(\d{4}).*(\d{4})$', '$1****...****$2'
    Update-SyncroBitLockerField -RecoveryKey $rk.RecoveryKey
    Write-Log ("BitLocker key retrieved (volumeType: {0}, created: {1}, id: {2}, masked: {3})" -f $rk.VolumeType, $rk.Created, $rk.RecoveryId, $masked)
    if (-not (Update-SyncroBitLockerField -RecoveryKey $rk.RecoveryKey)) {
        Write-Log "Syncro update skipped/failed; key retrieved successfully but not written to Syncro." "WARN"
    }

    Write-Log "Done."
    exit 0

} catch {
    $msg = "ERROR: $($_.Exception.Message)"
    Write-Error $msg
    Write-Log $msg "ERROR"
    exit 1
}
