# PS-Graph API User

PowerShell script to manage delegated OAuth2 permission grants (scopes) on the **Microsoft Graph Command Line Tools** enterprise app in Entra ID.

## Overview

The `Graph_api_User_add_remove.ps1` script targets the Microsoft Graph Command Line Tools service principal (`14d82eec-204b-4c2f-b7e8-296a70dab67e`) and lets you view, add, or remove delegated permission scopes without accidentally wiping existing grants.

## Requirements

- PowerShell 7+ (the GUI frontend requires PowerShell 7 / `pwsh`)
- `Microsoft.Graph` module (auto-installed from PSGallery if missing)
- An Entra ID account with `Application.Read.All`, `DelegatedPermissionGrant.ReadWrite.All`, and `User.Read.All` permissions

## Usage

```powershell
.\Graph_api_User_add_remove.ps1 -Action <View|Add|Remove> [-Scopes <scope,...>] [-ConsentType <AllPrincipals|Principal>] [-PrincipalId <ObjectId>]
```

### Parameters

| Parameter | Required | Description |
|---|---|---|
| `-Action` | Yes | `View`, `Add`, or `Remove` |
| `-Scopes` | For Add/Remove | Space- or comma-separated scope names |
| `-ConsentType` | No | `AllPrincipals` (default) or `Principal` |
| `-PrincipalId` | For Principal consent | Object ID (GUID) of the target user |
| `-WhatIf` | No | Preview changes without applying them |

When `-Action View` is combined with `-ConsentType Principal` and `-PrincipalId`, only the grant for that specific user is shown. Omit `-PrincipalId` to list all per-user grants.

### Examples

```powershell
# View all current grants
.\Graph_api_User_add_remove.ps1 -Action View

# Add scopes (admin consent for all users)
.\Graph_api_User_add_remove.ps1 -Action Add -Scopes "User.Read","Mail.Read"

# Remove a scope
.\Graph_api_User_add_remove.ps1 -Action Remove -Scopes "Mail.Read"

# Add a scope for a specific user
.\Graph_api_User_add_remove.ps1 -Action Add -Scopes "Calendars.Read" -ConsentType Principal -PrincipalId "<ObjectId>"

# Preview changes without applying (-WhatIf)
.\Graph_api_User_add_remove.ps1 -Action Add -Scopes "User.Read" -WhatIf
```


## Frontend (PowerShell GUI)

A dark-themed Windows Forms frontend is included in `Graph_api_User_frontend.ps1`.

```powershell
.\Graph_api_User_frontend.ps1
```

### GUI notes

- Windows-only (uses `System.Windows.Forms`); must be run with `pwsh` (PowerShell 7+)
- Calls `Graph_api_User_add_remove.ps1` under the hood — both files must be in the same directory
- **UPN field** — enter a user principal name (`user@domain.com`) and click **Lookup ID** to resolve it to an Entra ID Object ID; the resolved GUID is shown read-only and passed automatically to the backend
- Supports **Preview only (-WhatIf)** mode
- Status bar shows last run time, action, and exit code

## Notes

- Add and Remove operations **merge/subtract** from the existing scope list — no grants are wiped by accident.
- If all scopes are removed from a grant, the grant itself is deleted.
- Supports `-WhatIf` and `-Confirm` via `SupportsShouldProcess`.
