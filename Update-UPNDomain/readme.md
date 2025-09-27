# Bulk UPN Rename Script

This PowerShell script renames Microsoft 365 user accounts from  
`@coraswellness.org` to `@inperium.onmicrosoft.com`.

## Features
- Renames UPNs via Microsoft Graph
- Skips on-prem synced and guest accounts
- Optional: updates Exchange Online SMTP addresses
- Conflict check to prevent duplicate UPNs
- Supports `-WhatIf` dry run
- Logs results to CSV

## Requirements
- PowerShell 5.1 or 7+
- Modules:
  - Microsoft.Graph.Users
  - ExchangeOnlineManagement (only if using `-UpdateMailbox`)
- Target domain must be verified in your tenant

## Usage
```powershell
# Preview changes
.\Update-UpnDomain.ps1 -WhatIf

# Rename UPNs only
.\Update-UpnDomain.ps1

# Rename UPNs and update mailbox addresses
.\Update-UpnDomain.ps1 -UpdateMailbox

## Output
A CSV log is created with:
- OldUPN
- NewUPN
- Status
- Notes

### Example CSV
```csv
TimeStamp,DisplayName,OldUPN,NewUPN,AccountEnabled,OnPremSynced,Status,Notes
2025-09-26T17:40:12,John Smith,john.smith@coraswellness.org,john.smith@inperium.onmicrosoft.com,True,False,UPNUpdated,UPN changed via Graph.
2025-09-26T17:40:13,Jane Doe,jane.doe@coraswellness.org,jane.doe@inperium.onmicrosoft.com,True,True,Skipped,On-prem synced; change in AD.
2025-09-26T17:40:14,Bob Lee,bob.lee@coraswellness.org,bob.lee@inperium.onmicrosoft.com,False,False,Failed,Target UPN already exists: bob.lee@inperium.onmicrosoft.com
