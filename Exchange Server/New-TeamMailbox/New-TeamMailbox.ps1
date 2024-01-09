<#
  .SYNOPSIS
  Creates a new shared mailbox, security groups for full access and send-as permission
  and adds the security groups to the shared mailbox configuration.

  Thomas Stensitzki

  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

  Version 1.4, 2024-01-08

  Please send ideas, comments and suggestions to support@granikos.eu

  .LINK
  https://scripts.granikos.eu

  .DESCRIPTION
  This scripts creates a new shared mailbox (aka team mailbox) and security groups
  for full access and and send-as delegation. Security groups are created using a
  naming convention. The script can create a shared mailbox in Exchange Online uesing
  the -RemoteMailbox switch in combination with Entra ID Connect.

  Starting with v1.4 the script sets the sAMAccountName of the security groups to the
  group name to avoid numbered name extension of sAMAccountName.

  .NOTES
  Requirements
  - Windows Server 2016+
  - Exchange 2016+ Management Shell (aka EMS)
  - settings.xml in same folder as script containg general settings and group settings


  Revision History
  --------------------------------------------------------------------------------
  1.0 Initial community release
  1.1 Prefix seperator added, PowerShell hygiene
  1.2 PowerShell fixes, no functionality update
  1.3 Set extensionAttr14 to "AADsync" to make AD-objects (Mailbox, DistributionGroups) public to AzureAD
  1.4 Set sAMAccountName to group name for avoid numbered name extension of sAMAccountName

  .PARAMETER TeamMailboxName
  Name attribute of the new team mailbox

  .PARAMETER TeamMailboxDisplayName
  Display name attribute of the new team mailbox

  .PARAMETER TeamMailboxAlias
  Alias attribute of the new team mailbox

  .PARAMETER TeamMailboxSmtpAddress
  Primary SMTP address attribute the new team mailbox

  .PARAMETER DepartmentPrefix
  Department prefix for automatically generated security groups (optional)

  .PARAMETER GroupFullAccessMembers
  String array containing full access members

  .PARAMETER GroupFullAccessMembers
  String array containing send as members

  .PARAMETER RemoteMailboxFutureParameter
  Switch to create the shared mailbox in Exchange Online. Requires a hybrid configuration with Exchange Online.

  .EXAMPLE
  Create a new team mailbox, empty full access and empty send-as security groups

  .\New-TeamMailbox.ps1 -TeamMailboxName "TM-Exchange Admins" -TeamMailboxDisplayName "Exchange Admins" -TeamMailboxAlias "TM-ExchangeAdmins" -TeamMailboxSmtpAddress "ExchangeAdmins@mcsmemail.de" -DepartmentPrefix "IT"
#>

param (
    [parameter(Mandatory, HelpMessage = 'Team Mailbox Name')]
    [string]$TeamMailboxName,
    [parameter(Mandatory, HelpMessage = 'Team Mailbox Display Name')]
    [string]$TeamMailboxDisplayName,
    [parameter(Mandatory, HelpMessage = 'Team Mailbox Alias')]
    [string]$TeamMailboxAlias,
    [string]$TeamMailboxSmtpAddress = '',
    [string]$DepartmentPrefix = '',
    $GroupFullAccessMembers = @(''),
    $GroupSendAsMember = @()
)

# Script Path
$scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Path

if (Test-Path -Path ('{0}\Settings.xml' -f $scriptPath)) {

    # Load Script settings
    [xml]$Config = Get-Content -Path ('{0}\Settings.xml' -f $scriptPath)

    Write-Verbose -Message 'Loading script settings'

    # Group settings
    $groupPrefix = $Config.Settings.GroupSettings.Prefix
    $groupSendAsSuffix = $Config.Settings.GroupSettings.SendAsSuffix
    $groupFullAccessSuffix = $Config.Settings.GroupSettings.FullAccessSuffix
    $groupTargetOU = $Config.Settings.GroupSettings.TargetOU
    $groupDomain = $Config.Settings.GroupSettings.Domain
    $groupPrefixSeperator = $Config.Settings.GroupSettings.Seperator

    # Team mailbox settings
    $teamMailboxTargetOU = $Config.Settings.AccountSettings.TargetOU

    # General settings
    $sleepSeconds = $Config.Settings.GeneralSettings.Sleep
    $remoteRoutingAddress = $Config.Settings.GeneralSettings.RemoteRoutingAddress

    # Extensions settings
    $extAttr14 = $Config.Settings.Extensions.extAttr14

    Write-Verbose -Message 'Script settings loaded'
}
else {
    # Ooops, settings XML is missing
    Write-Error -Message 'Script settings file settings.xml missing'
    exit 99
}

# Add department prefix to group prefix, if configured
if ($DepartmentPrefix -ne '') {
    # TIP: Change pattern as needed
    $groupPrefix = ('{0}{1}{2}' -f $groupPrefix, $DepartmentPrefix, $groupPrefixSeperator)
}

# Create shared team mailbox
if ($RemoteMailbox) {
    # Create mailbox as remote mailbox
    Write-Verbose -Message ('New-RemoteMailbox -Shared -Name {0} -Alias {1} -RemoteRoutingAddress {1}@{2}' -f $TeamMailboxName, $TeamMailboxAlias, $remoteRoutingAddress)

    if ($TeamMailboxSmtpAddress -ne '') {
        # create new remote mailbox with dedicated email address
        $null = New-RemoteMailbox -Shared -Name $TeamMailboxName -Alias $TeamMailboxAlias -OrganizationalUnit $teamMailboxTargetOU -PrimarySmtpAddress $TeamMailboxSmtpAddress -DisplayName $TeamMailboxDisplayName -RemoteRoutingAddress ('{0}@{1}' -f $TeamMailboxAlias, $remoteRoutingAddress)
    }
    else {
        # create a new remote mailbox using email address policies
        $null = New-RemoteMailbox -Shared -Name $TeamMailboxName -Alias $TeamMailboxAlias -OrganizationalUnit $teamMailboxTargetOU -DisplayName $TeamMailboxDisplayName -RemoteRoutingAddress ('{0}@{1}' -f $TeamMailboxAlias, $remoteRoutingAddress)
    }

    # Copy sent messages into the shared mailbox
    # TODO: Must be set in EXO
    # Write-Verbose -Message ('Set-Mailbox -Identity  {0} -MessageCopyForSendOnBehalfEnabled:$true -MessageCopyForSentAsEnabled:$true' -f $TeamMailboxName)
    # $null = Set-Mailbox -Identity $TeamMailboxName -MessageCopyForSendOnBehalfEnabled:$true -MessageCopyForSentAsEnabled:$true

    # Add extensionAttribute (i.e., to include object in Entra ID Connect sync)
    # Change extension attribute number as needed
    # Value is read from settings.xml
    Write-Verbose -Message ('Set-Mailbox -Identity  {0} -CustomAttribute14 {1}' -f $TeamMailboxName, $extAttr14)
    $null = Set-RemoteMailbox -Identity $TeamMailboxName -CustomAttribute14 $extAttr14

}
else {
    # Create mailbox as on-premises mailbox
    Write-Verbose -Message ('New-Mailbox -Shared -Name {0} -Alias {1}' -f $TeamMailboxName, $TeamMailboxAlias)

    if ($TeamMailboxSmtpAddress -ne '') {
        # create new mailbox with dedicated email address
        $null = New-Mailbox -Shared -Name $TeamMailboxName -Alias $TeamMailboxAlias -OrganizationalUnit $teamMailboxTargetOU -PrimarySmtpAddress $TeamMailboxSmtpAddress -DisplayName $TeamMailboxDisplayName
    }
    else {
        # create a new mailbox using email address policies
        $null = New-Mailbox -Shared -Name $TeamMailboxName -Alias $TeamMailboxAlias -OrganizationalUnit $teamMailboxTargetOU -DisplayName $TeamMailboxDisplayName
    }

    # Copy sent messages into the shared mailbox
    Write-Verbose -Message ('Set-Mailbox -Identity  {0} -MessageCopyForSendOnBehalfEnabled:$true -MessageCopyForSentAsEnabled:$true' -f $TeamMailboxName)
    $null = Set-Mailbox -Identity $TeamMailboxName -MessageCopyForSendOnBehalfEnabled:$true -MessageCopyForSentAsEnabled:$true

    # Add extensionAttribute (i.e. to include object in Entra ID Connect sync)
    Write-Verbose -Message ('Set-Mailbox -Identity  {0} -CustomAttribute14 {1}' -f $TeamMailboxName, $extAttr14)
    $null = Set-Mailbox -Identity $TeamMailboxName -CustomAttribute14 $extAttr14
}

#region Create FullAccess ecurity group

# Create Full Access group for Team Mailbox
$groupName = ('{0}{1}{2}' -f $groupPrefix, $TeamMailboxAlias, $groupFullAccessSuffix)
$notes = ('FullAccess for mailbox: {0}' -f $TeamMailboxName)
$primaryEmail = ('{0}@{1}' -f $groupName, $groupDomain)

# 2024-01-08 - v1.4
# Set sAMAccountName to group name for avoid numbered name extension of sAMAccountName, change sAMAccountName pattern as needed
$groupSamAccountName = $groupName

Write-Host ('Creating new FullAccess Group: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

try {
    # Check ig security group already exists
    Write-Verbose -Message ('Get-ADGroup -Identity {0} -ErrorAction SilentlyContinue' -f $groupSamAccountName)
    $adGroupFound = Get-ADGroup -Identity $groupSamAccountName -ErrorAction SilentlyContinue
}
catch {
    # Group does not exist
    $adGroupFound = $false
}

# 2024-01-08 - v1.4
if ($adGroupFound -eq $false) {

    # No FullAccess group exists, create new group

    Write-Verbose -Message ('New-DistributionGroup -Name {0} -Type Security -OrganizationalUnit {1} -PrimarySmtpAddress {2}' -f $groupName, $groupTargetOU, $primaryEmail)

    if (($GroupFullAccessMembers | Measure-Object).Count -ne 0) {

        Write-Host ('Creating FullAccess group and adding members: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

        $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Members $GroupFullAccessMembers -Notes $notes -SamAccountName $groupSamAccountName

        Start-Sleep -Seconds $sleepSeconds

        # Hide FullAccess group from GAL
        Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true -EmailAddressPolicyEnabled $false

        # Add extensionAttribute (i.e., to include object in Entra ID Connect sync)
        Write-Verbose -Message ('Set-DistributionGroup -Identity  {0} -CustomAttribute14 {1}' -f $primaryEmail, $extAttr14)
        $null = Set-DistributionGroup -Identity $primaryEmail -CustomAttribute14 $extAttr14

    }
    else {

        Write-Host ('Creating empty FullAccess group: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

        $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Notes $notes -samAccountName $groupSamAccountName

        Start-Sleep -Seconds $sleepSeconds

        # Hide FullAccess group from GAL
        Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true

        # Add extensionAttribute (i.e., to include object in Entra ID Connect sync)
        Write-Verbose -Message ('Set-DistributionGroup -Identity  {0} -CustomAttribute14 {1}' -f $primaryEmail, $extAttr14)
        $null = Set-DistributionGroup -Identity $primaryEmail -CustomAttribute14 $extAttr14
    }

    # Add full access group to mailbox permissions

    Write-Verbose -Message ('Add-MailboxPermission -Identity {0} -User {1}' -f $TeamMailboxName, $primaryEmail)

    $null = Add-MailboxPermission -Identity $TeamMailboxName -User $primaryEmail -AccessRights FullAccess -InheritanceType all


}
else {
    Write-Host ('FullAccess Group with sAMAccountName {0} already exists. Script will exit.' -f $groupSamAccountName)
    exit 99
}

#endregion

#region Create SendAs security group

# Create Send As group
$groupName = ('{0}{1}{2}' -f $groupPrefix, $TeamMailboxAlias, $groupSendAsSuffix)
$notes = ('SendAs for mailbox: {0}' -f $TeamMailboxName)
$primaryEmail = ('{0}@{1}' -f $groupName, $groupDomain)

# 2024-01-08 - v1.4
# Set sAMAccountName to group name for avoid numbered name extension of sAMAccountName, change sAMAccountName pattern as needed
$groupSamAccountName = $groupName

Write-Host ('Creating new SendAs Group: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

try {
    # Check ig security group already exists
    Write-Verbose -Message ('Get-ADGroup -Identity {0} -ErrorAction SilentlyContinue' -f $groupSamAccountName)
    $adGroupFound = Get-ADGroup -Identity $groupSamAccountName -ErrorAction SilentlyContinue
}
catch {
    # Group does not exist
    $adGroupFound = $false
}

# 2024-01-08 - v1.4
if ($adGroupFound -eq $false) {

    Write-Verbose -Message ('New-DistributionGroup -Name {0} -Type Security -OrganizationalUnit {1} -PrimarySmtpAddress {2}' -f $groupName, $groupTargetOU, $primaryEmail)

    if (($GroupSendAsMember | Measure-Object).Count -ne 0) {

        Write-Host ('Creating SendAs group and adding members: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

        $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Members $GroupSendAsMember -Notes $notes -samAccountName $groupSamAccountName

        Start-Sleep -Seconds $sleepSeconds

        # Hide SendAs from GAL
        Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true

        # Add extensionAttribute (i.e., to include object in Entra ID Connect sync)
        Write-Verbose -Message ('Set-DistributionGroup -Identity  {0} -CustomAttribute14 {1}' -f $primaryEmail, $extAttr14)
        $null = Set-DistributionGroup -Identity $primaryEmail -CustomAttribute14 $extAttr14
    }
    else {

        Write-Host ('Creating empty SendAs group: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

        $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Notes $notes -samAccountName $groupSamAccountName

        Start-Sleep -Seconds $sleepSeconds

        # Hide SendAs from GAL
        Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true

        # Add extensionAttribute (i.e., to include object in Entra ID Connect sync)
        Write-Verbose -Message ('Set-DistributionGroup -Identity  {0} -CustomAttribute14 {1}' -f $primaryEmail, $extAttr14)
        $null = Set-DistributionGroup -Identity $primaryEmail -CustomAttribute14 $extAttr14
    }

    # Add SendAs permission
    Write-Verbose -Message ('Add-ADPermission -Identity {0} -User {1}' -f $TeamMailboxName, $groupName)

    $null = Add-ADPermission -Identity $TeamMailboxName -User $groupName -ExtendedRights 'Send-As'
}
else {
    Write-Host ('SendAs Group with sAMAccountName {0} already exists. Script will exit.' -f $groupSamAccountName)
    exit 99
}

Start-Sleep -Seconds $sleepSeconds

Write-Host ('Script finished. Team mailbox {0} created.' -f $TeamMailboxName)