#requires -version 5.1
<#
.SYNOPSIS
  Removes Azure AD *user* accounts from the local Administrators group, except for explicitly allowed identities.
.DESCRIPTION
  - Enumerates current members of the local "Administrators" group.
  - Identifies entries that are AzureAD *users* (not groups, not local/Microsoft accounts).
  - Skips the allowed identities (defaults include "AzureAD\CorasWellness" and "coraswellness@coraswellness.org").
  - Removes all other AzureAD user entries.
  - Emits a JSON summary to StdOut and logs to C:\ReduxTC\Logs\LocalAdminEnforcer\.
  - Supports -WhatIf and -Verbose.
.PARAMETER AllowedAzureADMembers
  Additional AzureAD identities to allow (case-insensitive). Accepts forms like "AzureAD\Name" or UPN "name@domain".
.PARAMETER LogRoot
  Root folder to write logs. Default: C:\ReduxTC\Logs\LocalAdminEnforcer
.PARAMETER JsonOnly
  If set, only the JSON summary is written to StdOut (logging still occurs).
.EXAMPLE
  .\Enforce-LocalAdmins-AzureAD.ps1 -WhatIf -Verbose
.EXAMPLE
  .\Enforce-LocalAdmins-AzureAD.ps1 -AllowedAzureADMembers 'AzureAD\AnotherUser'
.NOTES
  Version: 1.0.0
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string[]]$AllowedAzureADMembers = @(
        'AzureAD\CorasWellness',
        'AzureAD\coraswellness@coraswellness.org',
        'coraswellness@coraswellness.org'
    ),
    [string]$LogRoot = 'C:\ReduxTC\Logs\LocalAdminEnforcer',
    [switch]$JsonOnly
)

#region Helpers
$Script:ScriptVersion = '1.0.0'

function Test-IsAdmin {
    $current = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($current)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function New-LogFile {
    param(
        [string]$Root
    )
    try {
        if (-not (Test-Path -LiteralPath $Root)) {
            New-Item -ItemType Directory -Path $Root -Force | Out-Null
        }
        $stamp = (Get-Date).ToString('yyyy-MM-dd_HH-mm-ss')
        $file  = Join-Path $Root ("LocalAdminEnforcer_{0}.log" -f $stamp)
        New-Item -ItemType File -Path $file -Force | Out-Null
        return $file
    } catch {
        Write-Warning "Failed to create log directory or file at $Root. $_"
        return $null
    }
}

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if ($Script:LogFile) {
        try { Add-Content -Path $Script:LogFile -Value $line } catch {}
    }
    if ($Level -eq 'ERROR') { Write-Error $Message }
    elseif ($Level -eq 'WARN') { Write-Warning $Message }
    elseif ($Level -eq 'DEBUG') { Write-Verbose $Message }
    else { Write-Host $Message }
}

function Test-IsAzureAdMember {
    param($Member)
    # Member is the custom object with Name,ObjectClass,PrincipalSource,SID
    if ($Member.PrincipalSource -eq 'AzureAD') { return $true }
    if ($Member.Name -like 'AzureAD\*') { return $true }
    if ($Member.SID -and ($Member.SID.Value -like 'S-1-12-1-*')) { return $true }
    return $false
}

function Get-IdentityVariants {
    param([string]$Name)
    $set = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    [void]$set.Add($Name)
    if ($Name -match '^AzureAD\\(.+)$') {
        $leaf = $Matches[1]
        [void]$set.Add($leaf)
    }
    return $set
}
#endregion Helpers

#region Preflight
$summary = [ordered]@{
    Timestamp        = (Get-Date).ToString('s')
    ComputerName     = $env:COMPUTERNAME
    ScriptVersion    = $Script:ScriptVersion
    LogFile          = $null
    Allowed          = @($AllowedAzureADMembers)
    Admins           = @()
    AzureADUsersSeen = @()
    Removed          = @()
    Kept             = @()
    Errors           = @()
}

if (-not (Test-IsAdmin)) {
    $msg = 'This script must be run in an elevated PowerShell session.'
    $summary.Errors += $msg
    $summaryJson = $summary | ConvertTo-Json -Depth 5
    Write-Output $summaryJson
    exit 1
}

$Script:LogFile = New-LogFile -Root $LogRoot
$summary.LogFile = $Script:LogFile
Write-Log "Starting Local Administrator AzureAD cleanup. Version $Script:ScriptVersion" 'INFO'
#endregion Preflight

try {
    $admins = Get-LocalGroupMember -Group 'Administrators' -ErrorAction Stop | ForEach-Object {
        [PSCustomObject]@{
            Name            = $_.Name
            ObjectClass     = $_.ObjectClass
            PrincipalSource = $_.PrincipalSource
            SID             = $_.SID
            Raw             = $_
        }
    }
    $summary.Admins = $admins | Select-Object Name,ObjectClass,PrincipalSource,@{n='SID';e={$_.SID.Value}}
} catch {
    $msg = "Failed to enumerate Administrators group. $($_.Exception.Message)"
    Write-Log $msg 'ERROR'
    $summary.Errors += $msg
    $summaryJson = $summary | ConvertTo-Json -Depth 5
    Write-Output $summaryJson
    exit 1
}

# Build allowed set (case-insensitive)
$allowedSet = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
foreach ($a in $AllowedAzureADMembers) { if ($a) { [void]$allowedSet.Add($a) } }

# Filter to AzureAD *users* only
$aadUsers = @()
foreach ($m in $admins) {
    if ($m.ObjectClass -eq 'User' -and (Test-IsAzureAdMember -Member $m)) {
        $aadUsers += $m
    }
}
$summary.AzureADUsersSeen = $aadUsers | Select-Object Name,PrincipalSource,@{n='SID';e={$_.SID.Value}}

# Decide removals
$toRemove = @()
foreach ($m in $aadUsers) {
    $variants = Get-IdentityVariants -Name $m.Name
    $isAllowed = $false
    foreach ($v in $variants) {
        if ($allowedSet.Contains($v)) { $isAllowed = $true; break }
    }
    if ($isAllowed) {
        $summary.Kept += $m.Name
        Write-Log ("Keeping allowed AzureAD user in Administrators: {0}" -f $m.Name) 'DEBUG'
    } else {
        $toRemove += $m
    }
}

# Remove
foreach ($m in $toRemove) {
    $targetName = $m.Name
    if ($PSCmdlet.ShouldProcess("Administrators", "Remove member '$targetName'")) {
        try {
            Remove-LocalGroupMember -Group 'Administrators' -Member $m.Raw -ErrorAction Stop
            $summary.Removed += $targetName
            Write-Log ("Removed AzureAD user from Administrators: {0}" -f $targetName) 'INFO'
        } catch {
            Write-Log "Direct removal by object failed for $targetName, retrying by name..." 'WARN'
            try {
                Remove-LocalGroupMember -Group 'Administrators' -Member $targetName -ErrorAction Stop
                $summary.Removed += $targetName
                Write-Log ("Removed AzureAD user from Administrators by name: {0}" -f $targetName) 'INFO'
            } catch {
                $msg = "Failed to remove '$targetName'. $($_.Exception.Message)"
                Write-Log $msg 'ERROR'
                $summary.Errors += $msg
            }
        }
    } else {
        Write-Log ("Would remove AzureAD user from Administrators: {0}" -f $targetName) 'INFO'
    }
}

# Emit summary
$summaryJson = $summary | ConvertTo-Json -Depth 5

if (-not $JsonOnly) {
    Write-Host "---- Summary ----"
    Write-Host ("Removed: {0}" -f ([string]::Join(', ', $summary.Removed)))
    Write-Host ("Kept (allowed): {0}" -f ([string]::Join(', ', $summary.Kept)))
    Write-Host ("Log: {0}" -f $summary.LogFile)
    Write-Host "-----------------"
}
Write-Output $summaryJson

# Exit code: 0 on success (even if changes were made), 1 on preflight/critical error
if ($summary.Errors.Count -gt 0) {
    exit 1
} else {
    exit 0
}
