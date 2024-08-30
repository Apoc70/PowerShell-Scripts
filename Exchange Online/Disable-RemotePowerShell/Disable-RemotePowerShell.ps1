<#
    .SYNOPSIS

    Disable Exchange Remote PowerShell for all users, except for members of a selected security group.

    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.0, 2023-02-03

    Please post ideas, comments, and suggestions at GitHub.

    .LINK

    https://scripts.Granikos.eu

    .DESCRIPTION

    This script sets the user attribute RemotePowerShellEnabled to FALSE for all users.
    You can and should provide an Active Directory universal security group containing all user accounts
    that should be able to use Exchange Remote PowerShell, e.g., Admins, Service Account, etc.

    .NOTES

    Requirements
    - Windows Server 2016 or newer
    - Global function PowerShell library, found here: http://scripts.granikos.eu
    - AciveDirectory PowerShell module
    - Exchange 2013+ Management Shell

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial community release

    .PARAMETER DaysToKeep

    Number of days Exchange and IIS log files should be retained, default is 30 days

    .PARAMETER AllowRPSGroupName

    Name of the Active Directory security group containing all user accounts that must have Remote PowerShell enabled.
    Add all user accounts (administrators, users, service accounts) to that security group.

    .PARAMETER LogAllowedUsers

    Switch to write information about RPS allowed users to the log file.

    .PARAMETER LogDisabledUsers

    Switch to write information about users RBS disabled to the log file.

    .EXAMPLE

    Disable Exchange Remote PowerShell for all users using default settings, and do not write any user details to a log file.

    .\Disable-RemotePowerShell

#>


[CmdletBinding()]
Param(
  [int]$DaysToKeep = 30,
  [string]$AllowRPSGroupName = 'AllowRemotePS',
  [switch]$LogAllowedUsers,
  [switch]$LogDisabledUsers
)

# Some error variables
$ERR_OK = 0
$ERR_GLOBALFUNCTIONSMISSING = 1098
$ERR_NONELEVATEDMODE = 1099

# Load Exchange Management Shell PowerShell Snap-In
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

# Import Active Directory Module
Import-Module -Name ActiveDirectory

# Import GlobalFunctions
if($null -ne (Get-Module -Name GlobalFunctions -ListAvailable).Version) {
  Import-Module -Name GlobalFunctions
}
else {
  Write-Warning -Message 'Unable to load GlobalFunctions PowerShell module.'
  Write-Warning -Message 'Open an administrative PowerShell session and run Import-Module GlobalFunctions'
  Write-Warning -Message 'Please check http://bit.ly/GlobalFunctions for further instructions'
  exit $ERR_GLOBALFUNCTIONSMISSING
}

$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write('Script started')

# Check if we are running in elevated mode
# function (c) by Michel de Rooij
function script:Test-IsAdmin {
  $currentPrincipal = New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList ( [Security.Principal.WindowsIdentity]::GetCurrent() )

  If( $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
    return $true
  }
  else {
    return $false
  }
}

If (Test-IsAdmin) {
    # We are running in elevated mode. Let's continue.

    $logger.Write(('{1} started, keeping last {0} days of log files.' -f $DaysToKeep, $ScriptName))
    $logger.Purge()

    # Get all users with enabled Remote PowerShell
    $AllRPSUsers = Get-User -ResultSize Unlimited -Filter 'RemotePowerShellEnabled -eq $true' | select SamAccountName, RemotePowerShellEnabled

    # log number of users with RemotePowerShellEnabled enabled
    $logger.Write( ('Users with RemotePowerShellEnabled: {0}' -f ($AllRPSUsers | Measure-Object).Count ))

    # Get all users from AllowRPSGroupName

    $AllowedRPSUsers = Get-ADGroupMember -Identity $AllowRPSGroupName -Recursive | ForEach-Object { Get-User -Identity $_.SamAccountName | Select-Object SamAccountName, RemotePowerShellEnabled }

    # log number of users that are allowed to have RemotePowerShellEnabled
    $logger.Write( ('Number of allowed users: {0}' -f ($AllowedRPSUsers | Measure-Object).Count ))

    # Enable Remote PowerShell for allowed users
    foreach ($AllowedUser in $AllowedRPSUsers) {
        if ($AllowedUser.RemotePowerShellEnabled -eq $false) {
            if($LogAllowedUsers) {
              # write user information to log file
                $logger.Write( ('Setting RemotePowerShellEnabled to TRUE for: {0}' -f $AllowedUser.SamAccountName ))
            }

            # Set RemotePowerShellEnabled to TRUE for the user
            Set-User $AllowedUser.SamAccountName -RemotePowerShellEnabled $true
        }
    }

    # Disable Remote PowerShell for all users
    foreach ($User in $AllRPSUsers) {

        if ($AllowedRPSUsers.SamAccountName -notcontains $User.SamAccountName) {
            if($LogDisabledUsers) {
                # write user information to log file
                $logger.Write( ('Setting RemotePowerShellEnabled to FALSE for: {0}' -f $User.SamAccountName ))
            }

            # Set RemotePowerShellEnabled to FALSE for the user
            Set-User $User.SamAccountName -RemotePowerShellEnabled $false

        }
    }

    $logger.Write('Script finished')

    Return $ERR_OK
}
else {
  # Ooops, the admin did it again.
  Write-Warning -Message 'You must run the script in elevated mode. Please start the PowerShell-Session with administrative privileges.'

  Return $ERR_NONELEVATEDMODE
}