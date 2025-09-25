# Update-OfficeLocation.ps1

## Overview
**Update-OfficeLocation.ps1** helps you *find and fill blank* `OfficeLocation` values for Microsoft 365 (Entra ID) users.  
It connects to Microsoft Graph, lists users with empty `OfficeLocation`, and interactively lets you **Enter**, **Repeat last**, **Skip**, or **Quit** for each user.  
All actions are logged to a CSV file for auditing.

---

## Features
- Connects to Microsoft Graph using the lightweight `Microsoft.Graph.Users` module.
- Finds all users with a blank `OfficeLocation`.
- Interactive options per user:
  - **E** → Enter a location manually (uses default if provided).
  - **R** → Repeat the last location entered.
  - **S** → Skip the user.
  - **Q** → Quit the session.
- Supports `-DefaultLocation` for bulk updates.
- Supports `-WhatIf` dry-run mode.
- Generates a detailed results CSV.

---

## Requirements
- **PowerShell 5.1+** or **PowerShell 7+**
- **Permissions**: Account with rights to update users (e.g., User Administrator, Global Administrator).
- **Graph scope**: `User.ReadWrite.All`
- **Module**: `Microsoft.Graph.Users` (installed automatically if missing).

---

## Parameters
- `-DefaultLocation <string>`  
  Prefills the “Enter” prompt. If you press **Enter** with nothing typed, this value is used.

- `-WhatIf` *(switch)*  
  Runs in dry-run mode. No updates are made, but logs still record actions.

- `-ResultsPath <string>`  
  File path for the results CSV. Defaults to the script directory with a timestamped name.

---

## Usage Examples
```powershell
# Run interactively for blank OfficeLocations
.\Update-OfficeLocation.ps1

# Set a default location and use Enter/Repeat for fast bulk updates
.\Update-OfficeLocation.ps1 -DefaultLocation "Denver HQ"

# Dry-run mode (preview changes without updating)
.\Update-OfficeLocation.ps1 -WhatIf

# Specify custom log file path
.\Update-OfficeLocation.ps1 -ResultsPath "C:\Logs\OfficeLocation_Audit.csv"
