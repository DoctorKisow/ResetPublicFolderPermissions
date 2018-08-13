# Set-PublicFolderPermissions
A script to reset then re-apply public folder permissions.

# Description
This script will reset then reapply public folder security back to baseline with the "Default" and "Anonymous" permissions set to "None".

# Usage
    Set-MailboxPermissions -PublicFolder <public folder> -PublicFolderOwner <alias>
    Set-MailboxPermissions -PublicFolder <public folder> -PublicFolderOwner <alias> -DefaultPermission <permission> -AnonymousPermission <permission>
