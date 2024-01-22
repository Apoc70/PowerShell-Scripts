<#
  .SYNOPSIS
  Creates a new room mailbox, security groups for full access and send-as permission
  and adds the security groups to the room mailbox configuration.

  Thomas Stensitzki

  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

  Version 1.4, 2024-01-10

  Please send ideas, comments and suggestions to support@granikos.eu

  .LINK
  http://scripts.granikos.eu

  .DESCRIPTION
  This scripts creates a new room mailbox and additonal security groups
  for full access and and send-as delegation. The security groups are created
  using a confgurable naming convention.

  All required settings are stored in a separate settings.xml file

  .NOTES
  Requirements
    - Windows Server 2016+
    - Exchange 2016+ Management Shell (aka EMS)
    - settings.xml in same folder as script containg general settings and group settings


  Revision History
  --------------------------------------------------------------------------------
  1.0 Initial community release
  1.1 Some PowerShell hygiene, issue #2 closed
  1.2 CalendarBooking added, issue #1 closed
  1.3 Set extensionAttr14 to "AADsync" to make AD objects (Mailbox, DistributionGroups) public to Entra ID
  1.4 Set sAMAccountName to group name for avoid numbered name extension of sAMAccountName

  .PARAMETER RoomMailboxName
  Name attribute of the new team mailbox

  .PARAMETER RoomMailboxDisplayName
  Display name attribute of the new team mailbox

  .PARAMETER RoomMailboxAlias
  Alias attribute of the new team mailbox

  .PARAMETER RoomMailboxSmtpAddress
  Primary SMTP address attribute the new team mailbox

  .PARAMETER DepartmentPrefix
  Department prefix for automatically generated security groups (optional)

  .PARAMETER GroupFullAccessMembers
  String array containing full access members

  .PARAMATER RoomCapacity
  Capacity of the room, this value will show in the Outlook room list

  .PARAMETER RoomPhoneNumber
  Phone number of a phone located in the room, this value will show in the Outlook room list

  .PARAMETER RoomList
  Add the new room mailbox to this existing room list

  .PARAMETER AutoAccept
  Set room mailbox to automatically accept booking requests

  .PARAMETER Language
  Locale setting for calendar regional configuration language, e.g. de-DE, en-US

  .EXAMPLE
  Create a new room mailbox, empty full access and empty send-as security groups

  .\New-RoomMailbox.ps1 -RoomMailboxName "MB - Conference Room" -RoomMailboxDisplayName "Board Conference Room" -RoomMailboxAlias "MB-ConferenceRoom" -RoomMailboxSmtpAddress "ConferenceRoom@mcsmemail.de" -DepartmentPrefix "C"

  .EXAMPLE
  Create a new room mailbox, empty full access and empty send-as security groups, and add room to room list "Building 1"

  .\New-RoomMailbox.ps1 -RoomMailboxName "MB - Conference Room" -RoomMailboxDisplayName "Board Conference Room" -RoomMailboxAlias "MP-ConferencRoom" -RoomMailboxSmtpAddress "ConferenceRoom@mcsmemail.de" -DepartmentPrefix "C" -RoomList 'Building 1'

#>
param (
  [parameter(Mandatory, HelpMessage = 'Room Mailbox Name')]
  [string]$RoomMailboxName,
  [parameter(Mandatory, HelpMessage = 'Room Mailbox Display Name')]
  [string]$RoomMailboxDisplayName,
  [parameter(Mandatory, HelpMessage = 'Room Mailbox Alias')]
  [string]$RoomMailboxAlias,
  [string]$RoomMailboxSmtpAddress = '',
  [string]$DepartmentPrefix = '',
  [int]$RoomCapacity = 0,
  [string]$RoomPhoneNumber = '',
  [string]$RoomList = '',
  [switch]$AutoAccept,
  [String[]]$GroupFullAccessMembers = @(''),
  [String[]]$GroupSendAsMember = @(),
  [string]$Language = 'de-DE'
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
  $groupCalendarBookingSuffix = $Config.Settings.GroupSettings.CalendarBookingSuffix
  $groupTargetOU = $Config.Settings.GroupSettings.TargetOU
  $groupDomain = $Config.Settings.GroupSettings.Domain
  $groupPrefixSeperator = $Config.Settings.GroupSettings.Seperator

  # Team mailbox settings
  $roomMailboxTargetOU = $Config.Settings.AccountSettings.TargetOU

  # General settings
  $sleepSeconds = $Config.Settings.GeneralSettings.Sleep

  # Additional access security groups for all room mailboxes
  $groupStaticFullAccess = $Config.Settings.GeneralSettings.StaticFullAccessGroup
  $groupStaticSendAs = $Config.Settings.GeneralSettings.StaticSendAsGroup

  # Extensions settings
  # Adjust as needed, if you use a different extension attribute
  $extAttr14 = $Config.Settings.Extensions.extAttr14

  Write-Verbose -Message 'Script settings loaded'
}
else {
  Write-Error -Message 'Script settings file settings.xml missing'
  exit 99
}

function Test-AdGroup {
  param(
    [string]$AdGroupName = ''
  )

  try {
    # Check if security group already exists
    Write-Verbose -Message ('Get-ADGroup -Identity {0} -ErrorAction SilentlyContinue' -f $RoomList)
    $adGroupFound = Get-ADGroup -Identity $groupSamAccountName -ErrorAction SilentlyContinue
  }
  catch {
    # Group does not exist
    $adGroupFound = $false
  }

  return $adGroupFound
}

### MAIN

Write-Host ('Note: Script will pause for {0}s between steps' -f ($sleepSeconds))

# Add department prefix to group prefix, if configured
if ($DepartmentPrefix -ne '') {
  # Change pattern as needed
  $groupPrefix = ('{0}{1}{2}' -f $groupPrefix, $DepartmentPrefix, $groupPrefixSeperator)
}

# Create shared team mailbox
Write-Verbose -Message ('New-Mailbox -Room -Name {0} -Alias {1}' -f $RoomMailboxName, $RoomMailboxAlias)

if ($RoomMailboxSmtpAddress -ne '') {
  $null = New-Mailbox -Room -Name $RoomMailboxName -Alias $RoomMailboxAlias -OrganizationalUnit $roomMailboxTargetOU -PrimarySmtpAddress $RoomMailboxSmtpAddress -DisplayName $RoomMailboxDisplayName
}
else {
  $null = New-Mailbox -Room -Name $RoomMailboxName -Alias $RoomMailboxAlias -OrganizationalUnit $roomMailboxTargetOU -DisplayName $RoomMailboxDisplayName
}

# Add extensionAttribute (i.e. to include object in Entra ID Connect sync)
Write-Verbose -Message ('Set-Mailbox -Identity  {0} -CustomAttribute14 {1}' -f $RoomMailboxName, $extAttr14)

$null = Set-Mailbox -Identity $RoomMailboxName -CustomAttribute14 $extAttr14

# Set room capacity, if defined
if ($RoomCapacity -ne 0) {
  Start-Sleep -Seconds $sleepSeconds

  # Set room capacity
  Write-Verbose -Message ('Setting room mailbox capacity to {0}' -f ($RoomCapacity))

  Set-Mailbox -Identity $RoomMailboxAlias -ResourceCapacity $RoomCapacity
}

# Configure calendar processing to autoaccept
if ($AutoAccept) {
  Start-Sleep -Seconds $sleepSeconds

  Write-Verbose -Message 'Setting calendar processing to AutoAccept'

  Set-CalendarProcessing -Identity $RoomMailboxAlias -AutomateProcessing AutoAccept -AllowConflicts $false -MaximumConflictInstances 0 -ConflictPercentageAllowed 20
}

# Configure Language Regional Configuration
if ($Language -ne '') {

  Write-Verbose -Message ('Setting calendar regional configuration language to {0}' -f $Language)

  Set-MailboxRegionalConfiguration -Identity $RoomMailboxAlias -Language $Language
}

# Add to room list
if ($RoomList -ne '') {

  if (Test-AdGroup -AdGroupName $RoomList) {

    Start-Sleep -Seconds $sleepSeconds

    Write-Verbose -Message ('Adding mailbox to room list {0}' -f $RoomList)

    Add-DistributionGroupMember -Identity $RoomList -Member $RoomMailboxAlias
  }
  else {
    Write-Verbose -Message ('Room list {0} nor found' -f $RoomList)
  }
}

# Change phone number
if ($RoomPhoneNumber -ne '') {
  try {

    $SamAccountName = (Get-Mailbox -Identity $RoomMailboxAlias).SamAccountName
    $null = Get-ADUser -Identity $SamAccountName | Set-ADUser -OfficePhone $RoomPhoneNumber

    Write-Verbose -Message ('Room phone number set to {0}' -f ($RoomPhoneNumber))
  }
  catch {
    Write-Verbose -Message 'Phone number could not be set'
  }
}

#region Create FullAccess security group

# Create Full Access group
$groupName = ('{0}{1}{2}' -f $groupPrefix, $RoomMailboxAlias, $groupFullAccessSuffix)
$notes = ('FullAccess for mailbox: {0}' -f $RoomMailboxName)
$primaryEmail = ('{0}@{1}' -f $groupName, $groupDomain)

# 2024-01-10 - v1.4
# Set sAMAccountName to group name for avoid numbered name extension of sAMAccountName, change sAMAccountName pattern as needed
# This script simply uses the group name as sAMAccountName
$groupSamAccountName = $groupName

Write-Host ('Creating new FullAccess Group: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

# 2024-01-08 - v1.4
if (!(Test-AdGroup -AdGroupName $groupSamAccountName)) {
  # Group does not exist
  Write-Verbose -Message ('New-DistributionGroup -Name {0} -Type Security -OrganizationalUnit {1} -PrimarySmtpAddress {2}' -f $groupName, $groupTargetOU, $primaryEmail)

  if (($GroupFullAccessMembers | Measure-Object).Count -ne 0) {

    Write-Host ('Creating FullAccess group and adding members: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

    $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Members $GroupFullAccessMembers -Notes $notes -SamAccountName $groupSamAccountName

    Start-Sleep -Seconds $sleepSeconds
  }
  else {

    Write-Host ('Creating empty FullAccess group: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

    $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Notes $notes -SamAccountName $groupSamAccountName

    Start-Sleep -Seconds $sleepSeconds

  }

  # Hide FullAccess group from GAL
  Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true -EmailAddressPolicyEnabled $false

  # Add extensionAttribute (i.e., to include object in Entra ID Connect sync)
  Write-Verbose -Message ('Set-DistributionGroup -Identity  {0} -CustomAttribute14 {1}' -f $primaryEmail, $extAttr14)

  $null = Set-DistributionGroup -Identity $primaryEmail -CustomAttribute14 $extAttr14

  # Add full access group to mailbox permissions

  Write-Verbose -Message ('Add-MailboxPermission -Identity {0} -User {1}' -f $RoomMailboxName, $primaryEmail)

  $null = Add-MailboxPermission -Identity $RoomMailboxName -User $primaryEmail -AccessRights FullAccess -InheritanceType all

  # Assign an existing static full access group, if configured
  # This group can be used to grant full access to all room mailboxes for a room management department
  if ($groupStaticFullAccess -ne '') {
    if (Get-DistributionGroup -Identity $groupStaticFullAccess -ErrorAction SilentlyContinue) {
      # Only assign static full access group, if it exists
      $null = Add-MailboxPermission -Identity $RoomMailboxName -User $groupStaticFullAccess -AccessRights FullAccess -InheritanceType all
    }
  }
}
else {
  # Group exists
  Write-Host ('FullAccess Group with sAMAccountName {0} already exists. Script will exit.' -f $groupSamAccountName)
  exit 99
}

#end region

#region Create SendAs security group

# Create Send As group
$groupName = ('{0}{1}{2}' -f $groupPrefix, $RoomMailboxAlias, $groupSendAsSuffix)
$notes = ('SendAs for mailbox: {0}' -f $RoomMailboxName)
$primaryEmail = ('{0}@{1}' -f $groupName, $groupDomain)

# 2024-01-08 - v1.4
# Set sAMAccountName to group name for avoid numbered name extension of sAMAccountName, change sAMAccountName pattern as needed
$groupSamAccountName = $groupName

Write-Host ('Creating new SendAs Group: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

# 2024-01-08 - v1.4
if (!(Test-AdGroup -AdGroupName $groupSamAccountName)) {

  Write-Verbose -Message ('New-DistributionGroup -Name {0} -Type Security -OrganizationalUnit {1} -PrimarySmtpAddress {2}' -f $groupName, $groupTargetOU, $primaryEmail)

  if (($GroupSendAsMember | Measure-Object).Count -ne 0) {

    Write-Host ('Creating SendAs group and adding members: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

    $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Members $GroupSendAsMember -Notes $notes -SamAccountName $groupSamAccountName

    Start-Sleep -Seconds $sleepSeconds
  }
  else {

    Write-Host ('Creating empty SendAs group: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

    $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Notes $notes -SamAccountName $groupSamAccountName

    Start-Sleep -Seconds $sleepSeconds
  }

  # Hide SendAs from GAL
  Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true

  # Add extensionAttribute (i.e., to include object in Entra ID Connect sync)
  Write-Verbose -Message ('Set-DistributionGroup -Identity  {0} -CustomAttribute14 {1}' -f $primaryEmail, $extAttr14)

  $null = Set-DistributionGroup -Identity $primaryEmail -CustomAttribute14 $extAttr14

  # Add SendAs permission
  Write-Verbose -Message ('Add-ADPermission -Identity {0} -User {1}' -f $RoomMailboxName, $groupName)
  $null = Add-ADPermission -Identity $RoomMailboxName -User $groupName -ExtendedRights 'Send-As'

  # Assign an existing static send as group, if configured
  if ($groupStaticFullAccess -ne '') {

    if (Get-DistributionGroup -Identity $groupStaticSendAs -ErrorAction SilentlyContinue) {
      $null = Add-ADPermission -Identity $RoomMailboxName -User $groupStaticSendAs -ExtendedRights 'Send-As'
    }
  }
}
else {
  # Group exists
  Write-Host ('SendAs Group with sAMAccountName {0} already exists. Script will exit.' -f $groupSamAccountName)
  exit 99
}

#endregion

#region Create CalendarBooking security group

# Create CalendarBooking group
$groupName = ('{0}{1}{2}' -f $groupPrefix, $RoomMailboxAlias, $groupCalendarBookingSuffix)
$notes = ('CalendarBooking for mailbox: {0}' -f $RoomMailboxName)
$primaryEmail = ('{0}@{1}' -f $groupName, $groupDomain)

# 2024-01-08 - v1.4
# Set sAMAccountName to group name for avoid numbered name extension of sAMAccountName, change sAMAccountName pattern as needed
$groupSamAccountName = $groupName

Write-Host ('Creating new CalendarBooking Group: {0} (sAMAccountName: {1})' -f $groupName, $groupSamAccountName)

# 2024-01-08 - v1.4
if (!(Test-AdGroup -AdGroupName $adGroupFound)) {
  Write-Verbose -Message ('New-DistributionGroup -Name {0} -Type Security -OrganizationalUnit {1} -PrimarySmtpAddress {2}' -f $groupName, $groupTargetOU, $primaryEmail)

  $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Notes $notes -samAccountName $groupSamAccountName

  # Add extensionAttribute (i.e., to include object in Entra ID Connect sync)
  Write-Verbose -Message ('Set-DistributionGroup -Identity  {0} -CustomAttribute14 {1}' -f $primaryEmail, $extAttr14)

  $null = Set-DistributionGroup -Identity $primaryEmail -CustomAttribute14 $extAttr14
}
else {
  # Group exists
  Write-Host ('CalendarBooking Group with sAMAccountName {0} already exists. Script will exit.' -f $groupSamAccountName)
  exit 99
}

#endregion

Write-Host ('Script finished. Team mailbox {0} created.' -f $RoomMailboxName)