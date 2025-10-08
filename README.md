# Update-OnPremPublicFolderPermissions.ps1

## Overview

`Update-OnPremPublicFolderPermissions.ps1` is a PowerShell script designed for Exchange Server 2019 administrators to **set public folder client permissions** in bulk. The script allows you to specify one or more users, a target public folder, and the exact permission those users should have on that folder. It will ensure only the specified permission is active for the provided users—removing any other permissions they may have. Optionally, you can apply these changes recursively to all child folders.

## Features

- **Bulk Permission Update:** Specify multiple users and ensure they have only the desired permission on the selected folder(s).
- **Recursive Application:** Apply permission changes to all subfolders with the `-Recursive` switch.
- **Permission Reset:** Removes any existing permissions for specified users that do not match the one provided (or removes all permissions if `-Permission None` or not specified).
- **Idempotent:** Running the script multiple times results in the same permissions, avoiding duplicates.

## Parameters

- `-FolderPath` (required): The path to the public folder (e.g., `\Finance\Reports`).
- `-User` (required): One or more users (comma-separated list or multiple `-User` switches).
- `-Permission` (optional): The ONLY permission these users should have (e.g., `Editor`, `Owner`, `Reviewer`). If not specified or set to `None`, all permissions for these users will be removed.
- `-Recursive` (switch): If present, applies changes to all child public folders.

## Supported Permission Values

- `Owner`
- `PublishingEditor`
- `Editor`
- `PublishingAuthor`
- `Author`
- `NonEditingAuthor`
- `Reviewer`
- `Contributor`
- `None` (removes all permissions)

## Example Usage

**Set Alice as Editor on a single folder (removing any other rights):**
```powershell
.\Update-OnPremPublicFolderPermissions.ps1 -FolderPath "\Finance\Reports" -User alice@domain.com -Permission Editor
```

**Set Alice and Bob as Reviewers on a folder and all its subfolders:**
```powershell
.\Update-OnPremPublicFolderPermissions.ps1 -FolderPath "\Finance" -User alice@domain.com,bob@domain.com -Permission Reviewer -Recursive
```

**Remove all permissions for Alice on a folder:**
```powershell
.\Update-OnPremPublicFolderPermissions.ps1 -FolderPath "\Finance\Reports" -User alice@domain.com -Permission None
```

**Set Bob as Author on all subfolders under Finance:**
```powershell
.\Update-OnPremPublicFolderPermissions.ps1 -FolderPath "\Finance" -User bob@domain.com -Permission Author -Recursive
```

## Requirements

- **Exchange Server 2019** (with Public Folders)
- PowerShell session with appropriate Exchange management permissions

## Notes

- The script must be run in an Exchange Management Shell or a remote PowerShell session with access to Exchange cmdlets.
- Permission changes only affect the users specified—other users’ permissions on the same folder are not altered.
- If a user already has the specified permission, no action will be taken for that user/folder.

---
