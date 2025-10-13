# Update-OnPremPublicFolderPermissions.ps1

## Overview

This script is designed to set public folder client permissions in bulk with Exchange Server On-Premises. The script allows you to specify one or more users, a target public folder, and the exact permission those users should have on that folder. It will ensure only the specified permission is active for the provided users—removing any other permissions they may have. Optionally, you can apply these changes recursively to all child folders.

## Features

- **Bulk Permission Update:** Specify multiple users and ensure they have only the desired permission on the selected folder(s).
- **Recursive Application:** Apply permission changes to all subfolders with the `-Recursive` switch.
- **Permission Reset:** Removes any existing permissions for specified users that do not match the one provided (or removes all permissions if `-AccessRights None` or not specified).
- **Idempotent:** Running the script multiple times results in the same permissions, avoiding duplicates.

## Parameters

```powershell
Update-OnPremPublicFolderPermissions.ps1
    -FolderPath <String[]>
    -Users <String[]>
    -AccessRights <String[]>
    [-Recurse]
    [-WhatIf]
    [-Confirm]
    [<CommonParameters>]


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
.\Update-OnPremPublicFolderPermissions.ps1 -FolderPath "\Finance\Reports" -Users alice@domain.com -AccessRights Editor
```

**Set Alice and Bob as Reviewers on a folder and all its subfolders:**
```powershell
.\Update-OnPremPublicFolderPermissions.ps1 -FolderPath "\Finance" -Users alice@domain.com,bob@domain.com -AccessRights Reviewer -Recursive
```

**Remove all permissions for Alice on a folder:**
```powershell
.\Update-OnPremPublicFolderPermissions.ps1 -FolderPath "\Finance\Reports" -Users alice@domain.com -AccessRights None
```

**Set Bob as Author on all subfolders under Finance:**
```powershell
.\Update-OnPremPublicFolderPermissions.ps1 -FolderPath "\Finance" -Users bob@domain.com -AccessRights Author -Recursive
```

## Notes

- The script must be run in an Exchange Management Shell or a remote PowerShell session with access to Exchange cmdlets.
- Permission changes only affect the users specified — other users’ permissions on the same folder are not altered.
- If a user already has the specified permission, no action will be taken for that user/folder.

---
