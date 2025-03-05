<#
    .SYNOPSIS

    This script add, removes, or replaces public folder client permissions on selected public folders.

    .DESCRIPTION

    Long description

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    .NOTES

    Requirements

    - Windows Server 2019+
    - Exchange Server 2019+ Management Shell (aka EMS)

    Revision History
    --------------------------------------------------------------------------------
    1.0 Initial release


    .LINK

    https://somelink1.com/withmoreinformation

    .PARAMETER Action

    Action to add, remove, or replace permissions

    .PARAMETER PublicFolder

    Root public folder for changing permissions

    .PARAMETER User

    User to add, remove, or update

    .PARAMETER AccessRights

    Access rights to set for the user

    .PARAMETER Recurse

    Recurse through all subfolders

    .EXAMPLE

    .\Set-PublicFolderPermissions.ps1 -Action Add -PublicFolder '\RootFolder\SubFolder' -User 'JohnDoe' -AccessRights 'Reviewer'

    Add permissions for user JohnDoe to the public folder '\RootFolder\SubFolder' with the access rights 'Reviewer'

#>

# Parameter section with examples
# Additional information parameters: https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters
[CmdletBinding()]
param(
  [parameter(Mandatory, HelpMessage = 'Action to add, remove, or replace ', ParameterSetName = 'A')]
  [ValidateSet('Add', 'Remove', 'Replace', 'SetReadOnly')]
  [string]$Action,
  [parameter(Mandatory, HelpMessage = '(Root) public folder for changing permissions', ParameterSetName = 'A')]
  [string]$PublicFolder,
  [parameter(Mandatory, HelpMessage = 'User to add, remove, or update', ParameterSetName = 'A')]
  [string]$User,
  [ValidateSet('None', 'Owner', 'Reviewer')] # Needs update
  [string]$AccessRights,
  [switch]$Recurse,
  [ValidateSet('Anon', 'Default', 'AnonAndDefault')]
  [string]$PreserveD
)

# Variables
$scriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$scriptName = $MyInvocation.MyCommand.Name
$timeStamp = Get-Date -Format 'yyyy-MM-dd HHmmss'

# Global variable for storing public folders
$script:PublicFolders = $null

### functions

<#
    .DESCRIPTION

    Import required modules for the script to run
#>
function Import-RequiredModules {

  # Import central logging functions
  if ($null -ne (Get-Module -Name GlobalFunctions -ListAvailable).Version) {
    Import-Module -Name GlobalFunctions
  }
  else {
    Write-Warning -Message 'Unable to load GlobalFunctions PowerShell module.'
    Write-Warning -Message 'Please check http://bit.ly/GlobalFunctions for further instructions'
    exit
  }
}

<#
      .DESCRIPTION

      Get public folder hierarchy and store it in a global variable
#>
function Get-PublicFolderHierarchy {
  if ($Recurse) {
    Write-Verbose ('Gathering public folders recursively: {0}' -f $PublicFolder )
    $logger.Write( ('Gathering public folders recursively: {0}' -f $PublicFolder ) )
    $script:PublicFolders = Get-PublicFolder -Identity $PublicFolder -Recurse -ResultSize Unlimited -ErrorAction SilentlyContinue
    $logger.Write( ('Gathering public folders finished. Folders found: {0}' -f ($script:PublicFolders | Measure-Object).Count ) )
  }
  else {
    Write-Verbose ('Gathering public folder: {0}' -f $PublicFolder )
    $script:PublicFolders = Get-PublicFolder -Identity $PublicFolder -ErrorAction SilentlyContinue
  }
}

<#
  .DESCRIPTION

  This function replaces a single user permission with a permission in all public folders found in the global variable.
#>
function Replace-PublicFolderPermissions {
  [CmdletBinding()]
  param(
    [string]$UserToReplace,
    [string]$NewAccessRights
  )

  Get-PublicFolderHierarchy

  if ( ($script:PublicFolders | Measure-Object).Count -ne 0) {

    $i = 0
    $pfCount = ($script:PublicFolders | Measure-Object).Count

    foreach ($pf in $script:PublicFolders) {
      # progressbar
      Write-Progress -Activity 'Replacing Public Folder Permissions' -Status ('Processing folder {0}/{1}' -f $i, $pfCount) -PercentComplete (($i / $pfCount) * 100)

      $currentPermissions = Get-PublicFolderClientPermission ($pf.Identity).ToString() | Where-Object { ($_.User.DisplayName -eq $UserToReplace) }

      if ($currentPermissions) {
        # we found permissions - remove first
        $logger.Write( ('[{2}] REMOVING Access Rights: User {0} | {1}' -f $currentPermissions.User, $currentPermissions.AccessRights, ($pf.Identity).ToString() ) )
        $ret = Remove-PublicFolderClientPermission -Identity ($pf.Identity).ToString() -User $currentPermissions.User -Confirm:$False # -WhatIf
      }

      # add new access right for user to folder

      try {
        $logger.Write( ('[{2}] ADDING Access Rights: User {0} | {1}' -f $UserToAdd, $AccessRights, ($pf.Identity).ToString() ) )
        $ret = Add-PublicFolderClientPermission -Identity ($pf.Identity).ToString() -User $UserToReplace -AccessRights $AccessRights # -WhatIf
      }
      catch {}
      finally {}

      $i++
    }

  }
  else {
    # Nothing to do
    $logger.Write('No matching public folders found!')
  }

}

<#
  .DESCRIPTION
#>
function Add-PublicFolderPermissions {
  [CmdletBinding()]
  param(
    [string]$UserToAdd,
    [string]$NewAccessRights
  )

  Get-PublicFolderHierarchy

  if ( ($script:PublicFolders | Measure-Object).Count -ne 0) {

    $i = 0
    $pfCount = ($script:PublicFolders | Measure-Object).Count

    foreach ($pf in $script:PublicFolders) {
      # progressbar
      Write-Progress -Activity 'Adding Public Folder Permissions' -Status ('Processing folder {0}/{1}' -f $i, $pfCount) -PercentComplete (($i / $pfCount) * 100)

      $currentPermissions = Get-PublicFolderClientPermission ($pf.Identity).ToString() | Where-Object { ($_.User.DisplayName -eq $UserToAdd) }

      if ($currentPermissions) {
        # we found permissions - nothing to add
        $logger.Write( ('[{2}] User has permissions already: User {0} | {1}' -f $currentPermissions.User, $currentPermissions.AccessRights, ($pf.Identity).ToString() ) )
      }
      else {
        # add new access right for user to folder
        $logger.Write( ('[{2}] User has NO permissions | ADDING : User {0} | {1}' -f $UserToAdd, $AccessRights, ($pf.Identity).ToString() ) )

        try {
          $ret = Add-PublicFolderClientPermission -Identity ($pf.Identity).ToString() -User $UserToAdd -AccessRights $AccessRights #-WhatIf
        }
        catch {}
        finally {}
      }

      $i++
    }

  }
  else {
    # Nothing to do
    $logger.Write('No matching public folders found!')
  }
}

<#
  .DESCRIPTION

  This function removes a single user from all public folders found in the global variable.
#>
function Remove-PublicFolderPermissions {
  [CmdletBinding()]
  param(
    [string]$UserToRemove
  )

  Get-PublicFolderHierarchy

  if ( ($script:PublicFolders | Measure-Object).Count -ne 0) {

    $i = 0
    $pfCount = ($script:PublicFolders | Measure-Object).Count

    foreach ($pf in $script:PublicFolders) {
      # progressbar
      Write-Progress -Activity 'Removing Public Folder Permissions' -Status ('Processing folder {0}/{1}' -f $i, $pfCount) -PercentComplete (($i / $pfCount) * 100)

      $currentPermissions = Get-PublicFolderClientPermission ($pf.Identity).ToString() | Where-Object { ($_.User.DisplayName -eq $UserToRemove) }

      if ($currentPermissions) {
        # we found permissions - Let's remove

        $ret = Remove-PublicFolderClientPermission -Identity ($pf.Identity).ToString() -User $UserToRemove -WhatIf
        $logger.Write( ('[{2}] User Permission REMOVED: User {0} | {1}' -f $currentPermissions.User, $currentPermissions.AccessRights, ($pf.Identity).ToString() ) )

      }
      else {
        $logger.Write( ('[{12}] User has NO permissions. Nothing to remove. | ADDING : User {0} | {1}' -f $UserToAdd, ($pf.Identity).ToString() ) )
      }

    }
  }
  else {
    # Nothing to do
    $logger.Write('No matching public folders found!')
  }
}

<#
  .DESCRIPTION

  This function sets the public folder(s) gatherd in the global variable to READ-ONLY.

  All exisiting permissions are removed and replaced with READ-ONLY permissions.
  Only exception: ExchangePublicFolderManager
  Default, and Anonymous are removed completely.

#>
function Set-PublicFoldersReadOnly {

  Get-PublicFolderHierarchy

  if ( ($script:PublicFolders | Measure-Object).Count -ne 0) {

    $i = 0
    $pfCount = ($script:PublicFolders | Measure-Object).Count

    foreach ($pf in $script:PublicFolders) {
      # progressbar
      Write-Progress -Activity 'Setting Public Folder Permissions to READ-ONLY' -Status ('Processing folder {0}/{1}' -f $i, $pfCount) -PercentComplete (($i / $pfCount) * 100)

      $currentPermissions = Get-PublicFolderClientPermission ($pf.Identity).ToString() | Where-Object { $_.User -notlike 'ExchangePublicFolderManager' } # -and $_.User -notlike "Default" -and $_.User -notlike "Anonymous" }

      foreach ($permission in $currentPermissions) {

        # we found permissions - remove first
        $logger.Write( ('[{2}] REMOVING Access Rights: User {0} | {1}' -f $permission.User, [string]$permission.AccessRights, ($pf.Identity.ToString() ) ) )

        if ( ($permission.User -notlike "Default") -and ($permission.User -notlike "Anonymous") ) {

          if ( ($null -eq $permission.User.ADRecipient)) {
            $logger.Write( ('[{1}] User {0} NOT resolved in AD' -f $permission.User, ($pf.Identity).ToString() ) )
          }
          else {
            $ret = Remove-PublicFolderClientPermission -Identity ($pf.Identity).ToString() -User $permission.User.ToString() -Confirm:$False #-WhatIf
          }

        }
        else {
          # remove Anon and Default
          $ret = Remove-PublicFolderClientPermission -Identity ($pf.Identity).ToString() -User $permission.User.ToString() -Confirm:$False
        }

        try {

          if ( ($permission.User -notlike "Default") -and ($permission.User -notlike "Anonymous") ) {

            $logger.Write( ('[{2}] ADDING Access Rights: User {0} | {1}' -f $permission.User, $AccessRights, ($pf.Identity).ToString() ) )
            $ret = Add-PublicFolderClientPermission -Identity ($pf.Identity).ToString() -User $permission.User.ToString() -AccessRights $AccessRights # -WhatIf

          }
        }
        catch {}
        finally {}

      }
      $i++
    }

  }
  else {
    # Nothing to do
    $logger.Write('No matching public folders found!')
  }
}

### MAIN

# Import required module(s) first
Import-RequiredModules

# Create new logger
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
# Purge logs depening on LogFileRetention
$logger.Purge()
$logger.Write('Script started')

if ($Recurse) {
  $text = ('The script will {0} permissions to/from {1} and subfolders.' -f $Action, $PublicFolder)
  $logger.Write( $text )
}
else {
  $text = ('The script will {0} permissions to/from {1} only.' -f $Action, $PublicFolder)
  $logger.Write( $text )
}

if ($Action -ne 'SetReadOnly') {
  # try resolving user
  try {
    $resolvedUserDisplayName = (Get-Recipient $User).DisplayName
    $logger.Write( ('User {0} resolved to {1}' -f $User, $resolvedUserDisplayName) )
  }
  catch {
    $resolvedUserDisplayName = $null
    $logger.Write( ('Unable to resolve User {0}' -f $User) )

    throw ('Unable to resolve User {0}' -f $User)
  }

  Write-Host ("{0}`nUser: {1} `nPermission: {2}" -f $text, $User, $AccessRights)
}
else {
  Write-Host ('Setting folder(s) {0} to Read-Only' -f $PublicFolder)
}

switch ($Action) {
  'Add' { Add-PublicFolderPermissions -UserToAdd $resolvedUserDisplayName -NewAccessRights $AccessRights }
  'Replace' { Replace-PublicFolderPermissions -UserToReplace $resolvedUserDisplayName -NewAccessRights $AccessRights }
  'Remove' { Remove-PublicFolderPermissions }
  'SetReadOnly' {
    $AccessRights = 'Reviewer'
    Set-PublicFoldersReadOnly
  }
}

$logger.Write('Script finished')