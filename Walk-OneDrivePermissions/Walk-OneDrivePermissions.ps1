<# 
.SYNOPSIS
  Crawl all provisioned OneDrives and export a sharing/permissions report.

.REQUIREMENTS
  - Microsoft.Graph PowerShell SDK (Install-Module Microsoft.Graph -Scope AllUsers)
  - An Entra app registered with APPLICATION permissions granted admin consent:
      * Files.Read.All (or Sites.Read.All)
      * User.Read.All
  - Certificate-based auth (thumbprint on the machine running the script)

.PARAMETERS
  -TenantId            Your Entra tenant ID (GUID)
  -ClientId            The app registration (client) ID (GUID)
  -CertThumbprint      Thumbprint of the cert in CurrentUser/My or LocalMachine/My
  -OutputPath          Folder to write CSVs

.EXAMPLE
  .\OneDrive-Permissions-Audit.ps1 -TenantId "xxxx" -ClientId "yyyy" `
    -CertThumbprint "THUMB" -OutputPath "C:\Reports\OneDrivePerms"
#>

param(
  [Parameter(Mandatory=$false)][string]$TenantId,
  [Parameter(Mandatory=$false)][string]$ClientId,
  [Parameter(Mandatory=$false)][string]$CertThumbprint,
  [Parameter(Mandatory=$false)][string]$OutputPath,
  [switch]$StopOnError,
  [int]$MaxUsers = 0   # 0 = all
)

# ----- Utility: basic 429/5xx retry wrapper for Graph SDK calls -----
function Invoke-WithRetry {
  param(
    [Parameter(Mandatory=$false)][ScriptBlock]$Script,
    [int]$MaxAttempts = 6
  )
  $attempt = 0
  while ($true) {
    try {
      $attempt++
      return & $Script
    } catch {
      $ex = $_.Exception
      $msg = $ex.Message
      $status = $ex.ResponseStatusCode
      $retryAfter = 0
      if ($ex.ResponseHeaders.ContainsKey("Retry-After")) {
        [int]::TryParse($ex.ResponseHeaders["Retry-After"], [ref]$retryAfter) | Out-Null
      }
      if ($attempt -lt $MaxAttempts -and ($status -eq 429 -or ($status -ge 500 -and $status -lt 600))) {
        if ($retryAfter -le 0) { $retryAfter = [Math]::Min([int][Math]::Pow(2,$attempt), 60) }
        Write-Warning "Graph throttled or transient error ($status). Retry in $retryAfter s..."
        Start-Sleep -Seconds $retryAfter
      } else {
        throw
      }
    }
  }
}

# ----- Connect app-only (certificate) -----
Import-Module Microsoft.Graph.Users, Microsoft.Graph.Files, Microsoft.Graph.Sites -ErrorAction Stop

Connect-MgGraph -Scopes "User.Read.All, Sites.Read.All"
$context = Get-MgContext
Write-Host "Connected to $($context.TenantId) as app. Scopes: $($context.Scopes -join ', ')" -ForegroundColor Green

# ----- Prep output -----
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
$combinedCsv = Join-Path $OutputPath "OneDrive_Permissions_${timestamp}.csv"

$rows = New-Object System.Collections.Generic.List[object]

# ----- Get all users (you can add filters e.g. -Filter "accountEnabled eq true") -----
$allUsers = @()
$select = "id,userPrincipalName,displayName"
$batch = 999
$resp = Invoke-WithRetry { Get-MgUser -All -PageSize $batch -Property $select -ConsistencyLevel eventual -Count userCount -Select $select }
$allUsers = $resp

if ($MaxUsers -gt 0) {
  $allUsers = $allUsers | Select-Object -First $MaxUsers
}

Write-Host ("Found {0} users. Checking OneDrive provisioning..." -f $allUsers.Count)

# ----- Helper to enumerate drive items via delta -----
function Get-AllDriveItemsViaDelta {
  param([string]$DriveId)

  $items = @()
  $delta = $null
  do {
    $page = Invoke-WithRetry { Get-MgDriveItemDelta -DriveId $DriveId -Token $delta -ErrorAction Stop }
    if ($page.Value) { $items += $page.Value }
    $delta = $page.'@odata.deltaLink'  # when present, enumeration is complete
    if (-not $delta -and $page.'@odata.nextLink') {
      # next page uses 'token' query param; SDK’s -Token works with nextLink too
      $delta = $page.'@odata.nextLink'.Split('token=')[-1]
    }
  } while ($page.'@odata.nextLink' -or -not $delta)

  return $items
}

# ----- Crawl each user drive and collect permissions -----
$uIndex = 0
foreach ($u in $allUsers) {
  $uIndex++
  $upn = $u.UserPrincipalName
  try {
    Write-Host ("[{0}/{1}] {2} - checking drive..." -f $uIndex, $allUsers.Count, $upn)

    # Get user's OneDrive (if not provisioned, this will 404 in app-only)
    $drive = Invoke-WithRetry { Get-MgUserDrive -UserId $u.Id -ErrorAction Stop }

    if (-not $drive) {
      Write-Host "  No drive for $upn (not provisioned)." -ForegroundColor DarkYellow
      continue
    }

    Write-Host "  DriveId: $($drive.Id) - enumerating items (delta) ..."
    $items = Get-AllDriveItemsViaDelta -DriveId $drive.Id
    Write-Host ("  {0} items discovered." -f $items.Count)

    $i = 0
    foreach ($it in $items) {
      $i++
      # Skip deleted placeholders from delta
      if ($it.AdditionalProperties.ContainsKey("@microsoft.graph.deleted")) { continue }

      # Grab the path-like parentReference
      $path = $null
      if ($it.ParentReference.Path) {
        # path like: /drive/root:/Folder/Sub
        $path = $it.ParentReference.Path -replace '^/drive/root:',''
        if ([string]::IsNullOrEmpty($path)) { $path = "/" }
      } else { $path = "/" }

      # Fetch permissions for the item
      $perms = $null
      try {
        $perms = Invoke-WithRetry { Get-MgDriveItemPermission -DriveId $drive.Id -DriveItemId $it.Id -All -ErrorAction Stop }
      } catch {
        Write-Warning ("    Failed to read permissions on item {0}: {1}" -f $it.Id, $_.Exception.Message)
        if ($StopOnError) { throw }
        continue
      }

      foreach ($p in $perms) {
        # Permission resource can represent direct grantees or sharing links
        $isLink = $null -ne $p.Link
        $grantee  = $null
        $granteeId = $null
        if ($p.GrantedToV2.User) {
          $grantee   = $p.GrantedToV2.User.DisplayName
          $granteeId = $p.GrantedToV2.User.Email
        } elseif ($p.GrantedToV2.SiteUser) {
          $grantee   = $p.GrantedToV2.SiteUser.DisplayName
          $granteeId = $p.GrantedToV2.SiteUser.Email
        } elseif ($p.GrantedToV2.Group) {
          $grantee   = $p.GrantedToV2.Group.DisplayName
          $granteeId = $p.GrantedToV2.Group.Id
        } elseif ($p.GrantedTo) {
          # legacy facet
          $grantee   = $p.GrantedTo.User.DisplayName
          $granteeId = $p.GrantedTo.User.Email
        }

        $linkType  = $p.Link.LinkType
        $scope     = $p.Link.Scope
        $roles     = ($p.Roles -join ';')
        $expires   = $p.ExpirationDateTime
        $hasPassword = $p.Link.Password -ne $null
        $inheritedFrom = $p.InheritedFrom.Id

        $rows.Add([pscustomobject]@{
          ReportTimestamp   = (Get-Date)
          UserUPN           = $upn
          UserDisplayName   = $u.DisplayName
          DriveId           = $drive.Id
          ItemId            = $it.Id
          ItemName          = $it.Name
          ItemWebUrl        = $it.WebUrl
          ItemPath          = $path
          IsFolder          = [bool]($it.Folder -ne $null)
          PermissionId      = $p.Id
          PermissionKind    = if ($isLink) { "Link" } else { "Direct" }
          Roles             = $roles
          GrantedTo         = $grantee
          GrantedToId       = $granteeId
          LinkType          = $linkType
          LinkScope         = $scope     # anonymous, organization, or restricted
          LinkHasPassword   = $hasPassword
          Expires           = $expires
          InheritedFromItem = $inheritedFrom
        })
      }
    }

    # Per-user CSV (optional)
    $perUserCsv = Join-Path $OutputPath ("OneDrive_Perms_{0}_{1}.csv" -f ($upn -replace '[^a-zA-Z0-9@._-]','_'), $timestamp)
    $rows | Where-Object { $_.UserUPN -eq $upn } | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $perUserCsv
    Write-Host "  Wrote $perUserCsv"
  }
  catch {
    Write-Warning ("User {0} failed: {1}" -f $upn, $_.Exception.Message)
    if ($StopOnError) { throw }
  }
}

# Combined CSV
$rows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $combinedCsv
Write-Host "Combined report: $combinedCsv" -ForegroundColor Green
