<#
.SYNOPSIS
    A script to reset then re-apply public folder permissions.
.DESCRIPTION
    This script will reset then re-apply public folder security back to baseline with the 
    "Default" and "Anonymous" permissions set to "None".
.PARAMETER PublicFolder
    The name of the public folder.
.PARAMETER PublicFolderOwner
    The alias of the user or group you want to set as the "Publishing Editor".
.PARAMETER DefaultPermission
    The permission that you are setting the "Default" permission to.
    The options are Owner, Publishing Editor, Editor, Publishing Author,
    Author, Nonediting Author, Reviewer, Contributor and None.  The
    default permission is "None".
.PARAMETER AnonymousPermission
    The permission that you are setting the "Anonymous" permission to.
    The options are Owner, Publishing Editor, Editor, Publishing Author,
    Author, Nonediting Author, Reviewer, Contributor and None.  The
    default permission is "None".
.EXAMPLE
    C:\PS> .\Set-PublicFolderPermissions -PublicFolder "\Public Folder" -PublicFolderOwner "UserAlias" -DefaultPermission None -AnonymousPermission None
    C:\PS> .\Set-PublicFolderPermissions -PublicFolder "\Public Folder" -PublicFolderOwner "UserAlias" -DefaultPermission "Publishing Editor" -AnonymousPermission Author
    C:\PS> .\Set-PublicFolderPermissions -PublicFolder "\Public Folder" -PublicFolderOwner "UserAlias" 
.LINK
    https://github.com/DoctorKisow/ResetPublicFolderPermissions 
.NOTES
    Author: Matthew R. Kisow, D.Sc.
    Date:   August 13, 2018
#>

Param(
     [Parameter(Mandatory=$True, Position=0)]
     [string]$PublicFolder,
     [Parameter(Mandatory=$True, Position=1)]
     [string]$PublicFolderOwner,
     [Parameter(Mandatory=$True, Position=2)]
     [string]$DefaultPermission,
     [Parameter(Mandatory=$True, Position=3)]
     [string]$AnonymousPermission
)

Function ErrorChecking
{
    $IsPubFolder = Get-PublicFolder "$PublicFolder" -ErrorAction 'SilentlyContinue'
    if (-not $ISPubFolder)
    {
         Write-Host "The group $PublicFolder does not exist."
         exit 1
    }
    
    $IsUser = Get-Recipient -ANR $PublicFolderOwner -ErrorAction 'SilentlyContinue'
    if (-not $IsUser)
    {
         Write-Host "The owner $PublicFolderOwner does not exist or is not a mailbox."
         exit 1
    }

    if ("Owner", "Publishing Editor", "Editor", "Publishing Author", "Author", "Nonediting Author", "Reviewer", "Contributor", "None" -notcontains $DefaultPermission)
    {
         Write-Host "The Default permission $PublicFolderOwner is invalid."
         exit 1
    }

    if ("Owner", "Publishing Editor", "Editor", "Publishing Author", "Author", "Nonediting Author", "Reviewer", "Contributor", "None" -notcontains $AnonymousPermission)
    {
         Write-Host "The Default permission $AnonymousFolderOwner is invalid."
         exit 1
    }
}

Function VerifyUpdates
{
    # The de-facto "CONFIRM" and "ARE YOU SURE" function used to review and proceed with script execution.
    Write-Host "Are you sure that you want to update the security on this public folder."
    $CONTINUE = Read-Host "[Y]es or [N]o"

    while("Y","N" -notcontains $CONTINUE)
    {
        Write-Host "Incorrect response, please try again."
	    $CONTINUE = Read-Host "[Y]es or [N]o"
    }

    IF ($CONTINUE -eq 'N')
    {
        Write-Host "The public folder security updates have been aborted at the operators request."
        exit 1
    }
}

Function UpdatePublicFolderSecurity
{
    # Remove all permissions from the public folder.
    Get-PublicFolder -identity "$PublicFolder" -Recurse -ResultSize Unlimited | Get-PublicFolderClientPermission | Remove-PublicFolderClientPermission -Confirm:$false

    # Set the owner permissions to the Exchange public folder administrators group.
    Get-PublicFolder "$PublicFolder" -Recurse -ResultSize Unlimited | Add-PublicFolderClientPermission -User "SECExchangePublicFolderAdministrators" -AccessRights Owner

    # Obligatory and will not be needed once permissions have propogated.
    Get-PublicFolder "$PublicFolder" -Recurse -ResultSize Unlimited | Add-PublicFolderClientPermission -User "mkisow" -AccessRights Owner

    # Set the "Publishing Editor" permissions to the group supplied from the script.
    Get-PublicFolder "$PublicFolder" -Recurse -ResultSize Unlimited | Add-PublicFolderClientPermission -User "$PublicFolderOwner" -AccessRights "Publishing Editor"

    # Set the default user permissions.
    Get-PublicFolder "$PublicFolder" -Recurse -ResultSize Unlimited | Add-PublicFolderClientPermission -User Default -AccessRights $DefaultPermission

    # Set the anonymous user permissions.
    Get-PublicFolder "$PublicFolder" -Recurse -ResultSize Unlimited | Add-PublicFolderClientPermission -User Anonymous -AccessRights $AnonymousPermission
}

### Main Script

# If not already loaded, load the Exchange server PowerShell modules.
if ( (Get-PSSnapin -Name *Exchange* -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PSSnapin *Exchange*
}

Write-Host -ForegroundColor White "Set-PublicFolderPermissions"
Write-Host ""

ErrorChecking
VerifyUpdates
UpdatePublicFolderSecurity

# Exit Script
exit 0
