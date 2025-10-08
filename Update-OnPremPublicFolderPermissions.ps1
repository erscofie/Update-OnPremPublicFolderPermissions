<#
.SYNOPSIS
    Bulk update Public Folder permissions on Exchange Server 2019.

.DESCRIPTION
    This script sets the *exact* permissions for one or more users on a specified Public Folder:
    - Any existing permissions for the users not matching the specified Permission will be removed.
    - If Permission is "None" or not provided, all permissions for the users will be removed.
    - Optionally, changes can be applied recursively to all child folders.

.PARAMETER FolderPath
    The public folder path (e.g., \Finance\Reports).

.PARAMETER User
    One or more users (comma-separated or multiple -User switches).

.PARAMETER Permission
    The *only* permission(s) the users should have (e.g., Editor, Owner, Reviewer).
    If not specified or set to "None", all permissions for the users will be removed.

.PARAMETER Recursive
    If present, permissions are applied to all child public folders recursively.

.EXAMPLE
    .\Update-OnPremPublicFolderPermissions.ps1 -FolderPath "\Finance\Reports" -User alice@domain.com -Permission Editor

.EXAMPLE
    .\Update-OnPremPublicFolderPermissions.ps1 -FolderPath "\Finance" -User alice@domain.com,bob@domain.com -Permission Reviewer -Recursive

.EXAMPLE
    .\Update-OnPremPublicFolderPermissions.ps1 -FolderPath "\Finance\Reports" -User alice@domain.com -Permission None

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$FolderPath,

    [Parameter(Mandatory=$true)]
    [string[]]$User,

    [Parameter(Mandatory=$false)]
    [string]$Permission,

    [switch]$Recursive
)

function Set-ExactFolderPermission {
    param (
        [string]$FolderPath,
        [string[]]$User,
        [string]$Permission
    )

    foreach ($u in $User) {
        Write-Host "Processing folder: $FolderPath, User: $u"

        # Get current permission for the user
        $existing = Get-PublicFolderClientPermission -Identity $FolderPath -User $u -ErrorAction SilentlyContinue

        if (-not $Permission -or $Permission -eq 'None') {
            # Remove all permissions for this user
            if ($existing) {
                Remove-PublicFolderClientPermission -Identity $FolderPath -User $u -Confirm:$false
                Write-Host "Removed all permissions for $u on $FolderPath"
            } else {
                Write-Host "$u has no existing permissions on $FolderPath"
            }
        } else {
            if ($existing) {
                # Check if the current permission matches Permission
                $currentRights = $existing.AccessRights
                if ($currentRights -ne $Permission) {
                    # Remove all permissions and set the new one
                    Remove-PublicFolderClientPermission -Identity $FolderPath -User $u -Confirm:$false
                    Add-PublicFolderClientPermission -Identity $FolderPath -User $u -AccessRights $Permission
                    Write-Host "Reset permission to $Permission for $u on $FolderPath"
                } else {
                    Write-Host "$u already has $Permission on $FolderPath"
                }
            } else {
                # Add the permission directly
                Add-PublicFolderClientPermission -Identity $FolderPath -User $u -AccessRights $Permission
                Write-Host "Added $Permission for $u on $FolderPath"
            }
        }
    }
}

if ($Recursive) {
    Write-Host "Applying permissions recursively to $FolderPath and all child folders..."
    # Get all child folders including the root specified
    $allFolders = Get-PublicFolder -Identity $FolderPath -Recurse | Select-Object -ExpandProperty Identity
    foreach ($folder in $allFolders) {
        Set-ExactFolderPermission -FolderPath $folder -User $User -Permission $Permission
    }
} else {
    Set-ExactFolderPermission -FolderPath $FolderPath -User $User -Permission $Permission
}

Write-Host "Public Folder permissions update completed."