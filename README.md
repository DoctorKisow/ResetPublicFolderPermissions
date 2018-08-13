# Set-PublicFolderPermissions
An Exchange PowerShell script to remove and re-apply public folder permissions.

# Description
This script will reset then reapply public folder security back to baseline with the "Default" and "Anonymous" permissions set to "None".

# Usage
    Set-MailboxPermissions -PublicFolder <public folder> -PublicFolderOwner <alias> -DefaultPermission <permission> -AnonymousPermission <permission>
