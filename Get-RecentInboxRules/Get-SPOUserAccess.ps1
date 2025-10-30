<#
.SYNOPSIS
  Export SharePoint/OneDrive access activity from the Unified Audit Log for a given user.

.DESCRIPTION
  - Connects to EXO + Compliance (IPPSSession).
  - Searches the Unified Audit Log for SharePoint/OneDrive workloads.
  - Filters to common file-access operations (view, preview, download, modify, delete, move, upload, share).
  - Parses AuditData JSON and writes a tidy CSV.

.PARAMETER UserPrincipalName
  UPN of the user to investigate (required).

.PARAMETER DaysBack
  Lookback window in days (default: 7).

.PARAMETER SiteUrlContains
  Optional string filter: only include events where SiteUrl contains this text.

.PARAMETER OutputPath
  CSV output path (default: .\SPOD_Access_<UPN>_<yyyyMMddHHmm>.csv)

.EXAMPLE
  .\Get-SPODUserAccess.ps1 -UserPrincipalName user@domain.com -DaysBack 14

.EXAMPLE
  .\Get-SPODUserAccess.ps1 -UserPrincipalName user@domain.com -SiteUrlContains "teams/Finance"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$UserPrincipalName,

    [int]$DaysBack = 7,

    [string]$SiteUrlContains,

    [string]$OutputPath
)

function Write-Log {
    param(
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO',
        [string]$Message
    )
    $ts = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    Write-Host "[$ts] [$Level] $Message"
}

# Common SharePoint/OneDrive file operations to include
$Ops = @(
    # File views / access
    'FileAccessed','FilePreviewed','FileViewed','PageViewed','PreviewViewed',
    # Downloads / sync
    'FileDownloaded','FileSyncDownloadedFull','FileSyncDownloadedPartial',
    # Edits / moves / uploads / deletes
    'FileModified','FileMoved','FileUploaded','FileDeleted','FileRecycled','FileRestored',
    # Sharing / links
    'SharingLinkCreated','SharingSet','AnonymousLinkCreated','CompanyLinkCreated','SecureLinkCreated',
    'AddedToSecureLink','SharingInvitationCreated','SharingInvitationAccepted','SharingLinkDisabled'
)

# Build default output path
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $stamp = (Get-Date).ToString('yyyyMMddHHmm')
    $safeUpn = $UserPrincipalName.Replace('@','_')
    $OutputPath = ".\SPOD_Access_${safeUpn}_${stamp}.csv"
}

# Connect
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Log -Message "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Log -Message "Connecting to Compliance/Audit (IPPSSession)..."
    Connect-IPPSSession -ErrorAction Stop | Out-Null
}
catch {
    Write-Log -Level ERROR -Message "Failed to connect: $($_.Exception.Message)"
    throw
}

# Helper: chunked UAL search (avoids 5k cap)
function Search-UalChunked {
    param(
        [datetime]$Start,
        [datetime]$End,
        [string[]]$Operations,
        [string]$UserId
    )

    $results = @()
    # Use 24h slices to reduce result truncation; adjust if needed
    $cursor = $Start
    while ($cursor -lt $End) {
        $sliceEnd = $cursor.AddDays(1)
        if ($sliceEnd -gt $End) { $sliceEnd = $End }

        Write-Log -Message ("UAL slice: {0} -> {1}" -f $cursor.ToString("u"), $sliceEnd.ToString("u"))

        # Use SessionId/ReturnLargeSet for better paging
        $sid = [guid]::NewGuid().ToString()
        try {
            $chunk = Search-UnifiedAuditLog -StartDate $cursor -EndDate $sliceEnd `
                -UserIds $UserId -Operations $Operations -ResultSize 5000 `
                -SessionId $sid -SessionCommand ReturnLargeSet -ErrorAction SilentlyContinue
            if ($chunk) { $results += $chunk }
        }
        catch {
            Write-Log -Level WARN -Message "UAL query failed for slice: $($_.Exception.Message)"
        }

        $cursor = $sliceEnd
    }
    return $results
}

$endTime   = Get-Date
$startTime = (Get-Date).AddDays(-1 * $DaysBack)

Write-Log -Message "Querying UAL for $UserPrincipalName from $startTime to $endTime ..."
$raw = Search-UalChunked -Start $startTime -End $endTime -Operations $Ops -UserId $UserPrincipalName

if (-not $raw -or $raw.Count -eq 0) {
    Write-Log -Level WARN -Message "No SharePoint/OneDrive events found for user in the last $DaysBack day(s)."
}

# Parse and normalize rows
$rows = @()

foreach ($rec in $raw) {
    $data = $null
    try { $data = $rec.AuditData | ConvertFrom-Json } catch { $data = $null }

    # Pull common fields safely
    $workload     = $rec.Workload
    $operation    = $rec.Operations
    $timeUtc      = [datetime]::SpecifyKind($rec.CreationDate, [System.DateTimeKind]::Utc)
    $siteUrl      = if ($data) { $data.SiteUrl } else { $null }
    $sourceUrl    = if ($data) { $data.SourceRelativeUrl } else { $null }
    $objectId     = if ($data) { $data.ObjectId } else { $null }
    $fileName     = if ($data) { $data.SourceFileName } else { $null }
    $fileExt      = if ($data) { $data.SourceFileExtension } else { $null }
    $userAgent    = if ($data) { $data.UserAgent } else { $null }
    $clientIp     = if ($data) { $data.ClientIP } else { $null }
    $actorUpn     = if ($data) { $data.UserId } else { $rec.UserIds }
    $destUrl      = if ($data) { $data.DestinationRelativeUrl } else { $null }
    $eventSource  = if ($data) { $data.EventSource } else { $null }
    $siteId       = if ($data) { $data.Site } else { $null }
    $webId        = if ($data) { $data.WebId } else { $null }
    $listId       = if ($data) { $data.ListId } else { $null }
    $itemId       = if ($data) { $data.ListItemUniqueId } else { $null }
    $sharingType  = if ($data) { $data.SharingType } else { $null }
    $linkType     = if ($data) { $data.LinkType } else { $null }
    $recipient    = if ($data) { $data.TargetUserOrGroupName } else { $null }

    # Optional SiteUrl text filter
    $include = $true
    if (-not [string]::IsNullOrWhiteSpace($SiteUrlContains)) {
        $include = ($siteUrl -like ("*" + $SiteUrlContains + "*"))
    }

    if ($include) {
        $rows += [pscustomobject]@{
            TimeUTC      = $timeUtc
            Workload     = $workload
            Operation    = $operation
            ActorUPN     = $actorUpn
            ClientIP     = $clientIp
            UserAgent    = $userAgent
            SiteUrl      = $siteUrl
            ObjectId     = $objectId
            SourceUrl    = $sourceUrl
            DestinationUrl = $destUrl
            FileName     = $fileName
            FileExtension= $fileExt
            SiteId       = $siteId
            WebId        = $webId
            ListId       = $listId
            ItemId       = $itemId
            SharingType  = $sharingType
            LinkType     = $linkType
            TargetUserOrGroup = $recipient
            EventSource  = $eventSource
        }
    }
}

# Output
try {
    $rows | Sort-Object TimeUTC -Descending | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "Report written: $OutputPath"
}
catch {
    Write-Log -Level ERROR -Message "Failed to write CSV: $($_.Exception.Message)"
    throw
}

# Console preview
if ($rows.Count -gt 0) {
    Write-Host ""
    Write-Log -Message "Sample of results:"
    $rows | Select-Object TimeUTC, Operation, SiteUrl, FileName, ActorUPN, ClientIP | Format-Table -AutoSize
} else {
    Write-Log -Message "No results to display."
}
