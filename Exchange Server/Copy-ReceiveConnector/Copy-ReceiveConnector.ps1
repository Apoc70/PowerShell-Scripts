<#
    .SYNOPSIS
    Copy a selected receive connector and its configuration and permissions to other Exchange Servers

    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.9, 2026-10-27

    Please submit ideas, comments, and suggestions using GitHub.

    .LINK
    https://scripts.granikos.eu

    .DESCRIPTION
    This script copies a receive connector from a source Exchange Server to a single target Exchange server or to all Exchange servers.
    Configured permissions are copied as well, if required.

    All values read from the source connector are converted explicitly to strings before they are handed over to
    New-ReceiveConnector, Set-ReceiveConnector, and Add-ADPermission. This is required because the Exchange Management Shell
    uses PowerShell remoting and returns deserialized objects. Since PowerShell serialization payload signing is enabled by default,
    Exchange no longer converts deserialized Exchange types (e.g. IPRange) back automatically.

    .NOTES
    Requirements
    - Windows Server 2016, 2019, 2022, or 2025
    - Exchange Server 2016, Exchange Server 2019, or Exchange Server Subscription Edition (SE)
    - Exchange Management Shell

    Revision History
    --------------------------------------------------------------------------------
    1.0 Initial community release
    1.1 Domain Controller parameter added, permissions group copy added
    1.2 Move to FrontendTransport added, optional permission copy added, reset bindings added
    1.3 Update receive connector, if receive connector exists
    1.4 Fix to handle connector updates properly
    1.41 Minor fixes and update for Exchange 2016
    1.5 Issue #2 fixed
    1.6 Issue #3 fixed
    1.6.1 Minor fixes and tested with Exchange Server 2019
    1.7 UpdateExistingConnector added to bypass update confirmation for automation
    1.8 Check for Exchange Server 2019 and PowerShell Serialization Signing
    1.9 Issue #7 fixed (RemoteIPRanges and other values cannot be converted)
        All source values are converted explicitly, independent of serialization payload signing
        Source connector is read once, DomainController is used consistently
        Existing non-inherited AD permissions on the target connector are skipped
        Warning for server specific network bindings
        Additional connector properties copied, if supported by the target cmdlet
        -WhatIf support, SecondsToWait parameter added
        Unused ActiveDirectory module dependency removed

    .PARAMETER SourceServer
    Source Exchange server hosting the receive connector to copy.

    .PARAMETER ConnectorName
    Name of the receive connector to copy.

    .PARAMETER TargetServer
    Target Exchange server to copy the selected receive connector to.

    .PARAMETER DomainController
    The name of a domain controller. All read and write operations use this domain controller.

    .PARAMETER CopyToAllOther
    Switch to copy to all other Exchange servers with the Mailbox role.

    .PARAMETER CopyPermissions
    Copy non inherited source receive AD permissions to target receive connector. Inherited permissions will not be copied.
    Permissions that already exist on the target connector are skipped.

    .PARAMETER MoveToFrontend
    Change source connector transport role to FrontendTransport. This is required when you copy a receive connector from Exchange 2010 to Exchange 2013+.

    .PARAMETER ResetBindings
    Do not copy bindings but reset receive connector network bindings to 0.0.0.0:25.

    .PARAMETER UpdateExistingConnector
    Switch to force updating an existing receive connector on the target server without confirmation prompt.

    .PARAMETER ViewEntireForest
    Set scope to view entire Active Directory forest.

    .PARAMETER SecondsToWait
    Seconds to wait for Active Directory before AD permissions or permission groups are updated on a newly created connector. Default: 60

    .EXAMPLE
    Copy receive connector MYRECEIVECONNECTOR from server MBX01 to server MBX2

    .\Copy-ReceiveConnector.ps1 -SourceServer MBX01 -ConnectorName MYRECEIVECONNECTOR -TargetServer MBX2 -DomainController MYDC1.mcsmemail.de

    .EXAMPLE
    Copy receive connector MYRECEIVECONNECTOR from server MBX01 to all other Exchange servers

    .\Copy-ReceiveConnector.ps1 -SourceServer MBX01 -ConnectorName MYRECEIVECONNECTOR -CopyToAllOther -DomainController MYDC1.mcsmemail.de

    .EXAMPLE
    Copy receive connector MYRECEIVECONNECTOR from Exchange 2010 server MBX2010 to Exchange 2016 server MBX01, make it a FrontEnd-Connector, and reset network bindings

    .\Copy-ReceiveConnector.ps1 -SourceServer MBX2010 -ConnectorName MYRECEIVECONNECTOR -TargetServer MBX01 -MoveToFrontend -ResetBindings -DomainController MYDC1.mcsmemail.de

    .EXAMPLE
    Copy receive connector MYRECEIVECONNECTOR from server MBX01 to all other Exchange servers, including permissions, and update the connector if it already exists

    .\Copy-ReceiveConnector.ps1 -SourceServer MBX01 -ConnectorName MYRECEIVECONNECTOR -CopyToAllOther -DomainController MYDC1.mcsmemail.de -CopyPermissions -UpdateExistingConnector

    .EXAMPLE
    Show what would happen without changing anything

    .\Copy-ReceiveConnector.ps1 -SourceServer MBX01 -ConnectorName MYRECEIVECONNECTOR -TargetServer MBX2 -DomainController MYDC1.mcsmemail.de -CopyPermissions -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
  [parameter(Mandatory,HelpMessage='Source Exchange server to copy from')]
  [string]$SourceServer,
  [parameter(Mandatory,HelpMessage='Name of the source receive connector')]
  [string]$ConnectorName,
  [string]$TargetServer = '',
  [parameter(Mandatory,HelpMessage='Domain Controller name')]
  [string]$DomainController,
  [switch]$CopyToAllOther,
  [switch]$CopyPermissions,
  [switch]$MoveToFrontend,
  [switch]$ResetBindings,
  [switch]$UpdateExistingConnector,
  [switch]$ViewEntireForest,
  [ValidateRange(0,600)]
  [int]$SecondsToWait = 60
)

### FUNCTIONS -----------------------------

function Request-Choice {
  [CmdletBinding()]
  param(
    [string]$Caption = 'Really?'
  )
  $choices = [System.Management.Automation.Host.ChoiceDescription[]]@('&Yes','&No')

  [int]$defaultChoice = 1

  $choiceReturn = $Host.UI.PromptForChoice($Caption, '', $choices, $defaultChoice)

  return $choiceReturn
}

function ConvertTo-StringArray {
  # Converts single values, collections, and deserialized Exchange objects to a plain string array
  [CmdletBinding()]
  param(
    $Value
  )

  $result = New-Object -TypeName System.Collections.Generic.List[string]

  foreach($item in @($Value)) {
    if($null -ne $item) {
      $itemString = ([string]$item).Trim()
      if($itemString -ne '') {
        $result.Add($itemString)
      }
    }
  }

  return ,$result.ToArray()
}

function ConvertTo-Boolean {
  [CmdletBinding()]
  param(
    $Value
  )

  if($null -eq $Value) {
    return $false
  }

  return [System.Convert]::ToBoolean([string]$Value)
}

function ConvertTo-SizeString {
  # Converts a (deserialized) ByteQuantifiedSize like '36 MB (37,748,736 bytes)' to '37748736B'
  [CmdletBinding()]
  param(
    $Value
  )

  if($null -eq $Value) {
    return $null
  }

  $sizeString = ([string]$Value).Trim()

  if($sizeString -match '\(([\d\.,\s]+)\s*bytes\)') {
    return ('{0}B' -f ($Matches[1] -replace '[^\d]',''))
  }

  return $sizeString
}

function ConvertTo-FlagString {
  # Converts a (deserialized) flags enum like PermissionGroups or AuthMechanism to a comma separated string
  [CmdletBinding()]
  param(
    $Value,
    [string[]]$Exclude = @()
  )

  $flags = @((ConvertTo-StringArray -Value $Value) -join ',' -split ',' |
      ForEach-Object -Process { $_.Trim() } |
  Where-Object -FilterScript { ($_ -ne '') -and ($Exclude -notcontains $_) })

  return ($flags -join ', ')
}

function Test-ServerSpecificBinding {
  # Returns $true, if a binding uses a dedicated IP address instead of all IPv4/IPv6 addresses
  [CmdletBinding()]
  param(
    [string]$Binding
  )

  $address = $Binding.Substring(0, $Binding.LastIndexOf(':'))

  return (@('0.0.0.0','[::]','::') -notcontains $address)
}

function Get-ReceiveConnectorParameter {
  # Builds the splatting hashtable for New-ReceiveConnector or Set-ReceiveConnector from the source connector
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    $Connector,
    [Parameter(Mandatory)]
    [string]$TargetFqdn,
    [Parameter(Mandatory)]
    [ValidateSet('New-ReceiveConnector','Set-ReceiveConnector')]
    [string]$CommandName,
    [switch]$IsUpdate
  )

  $params = @{}

  # Multi-valued properties, the cause of issue #7
  $params.RemoteIPRanges = ConvertTo-StringArray -Value $Connector.RemoteIPRanges

  # Flags. The Custom permission group cannot be set and is represented by non-inherited AD permissions
  $permissionGroups = ConvertTo-FlagString -Value $Connector.PermissionGroups -Exclude 'Custom'
  if($permissionGroups -eq '') {
    $permissionGroups = 'None'
  }
  $params.PermissionGroups = $permissionGroups
  $params.AuthMechanism = ConvertTo-FlagString -Value $Connector.AuthMechanism

  $params.Fqdn = $TargetFqdn

  # Boolean properties
  $booleanProperties = @(
    'AdvertiseClientSettings'
    'BinaryMimeEnabled'
    'ChunkingEnabled'
    'DeliveryStatusNotificationEnabled'
    'DomainSecureEnabled'
    'EightBitMimeEnabled'
    'EnableAuthGSSAPI'
    'Enabled'
    'EnhancedStatusCodesEnabled'
    'LongAddressesEnabled'
    'OrarEnabled'
    'PipeliningEnabled'
    'RequireEHLODomain'
    'RequireTLS'
    'SuppressXAnonymousTls'
  )

  foreach($property in $booleanProperties) {
    if($null -ne $Connector.$property) {
      $params.$property = ConvertTo-Boolean -Value $Connector.$property
    }
  }

  # Properties that are handed over as string (TimeSpan, Unlimited<int>, int, enums, domain names)
  $stringProperties = @(
    'ConnectionInactivityTimeout'
    'ConnectionTimeout'
    'DefaultDomain'
    'ExtendedProtectionPolicy'
    'MaxAcknowledgementDelay'
    'MaxHopCount'
    'MaxInboundConnection'
    'MaxInboundConnectionPercentagePerSource'
    'MaxInboundConnectionPerSource'
    'MaxLocalHopCount'
    'MaxLogonFailures'
    'MaxProtocolErrors'
    'MaxRecipientsPerMessage'
    'MessageRateLimit'
    'MessageRateSource'
    'ProtocolLoggingLevel'
    'ServiceDiscoveryFqdn'
    'SizeEnabled'
    'TarpitInterval'
  )

  foreach($property in $stringProperties) {
    $propertyValue = [string]$Connector.$property
    if($propertyValue -ne '') {
      $params.$property = $propertyValue
    }
  }

  # ByteQuantifiedSize properties
  foreach($property in @('MaxHeaderSize','MaxMessageSize')) {
    $propertyValue = ConvertTo-SizeString -Value $Connector.$property
    if(($null -ne $propertyValue) -and ($propertyValue -ne '')) {
      $params.$property = $propertyValue
    }
  }

  # Properties that can be empty. Empty values are used to clear the property when updating an existing connector
  $banner = [string]$Connector.Banner
  $comment = [string]$Connector.Comment
  $tlsCertificateName = [string]$Connector.TlsCertificateName
  $tlsDomainCapabilities = ConvertTo-StringArray -Value $Connector.TlsDomainCapabilities

  if($banner -ne '') { $params.Banner = $banner } elseif($IsUpdate) { $params.Banner = $null }
  if($comment -ne '') { $params.Comment = $comment } elseif($IsUpdate) { $params.Comment = $null }
  if($tlsCertificateName -ne '') { $params.TlsCertificateName = $tlsCertificateName } elseif($IsUpdate) { $params.TlsCertificateName = $null }
  if($tlsDomainCapabilities.Count -gt 0) { $params.TlsDomainCapabilities = $tlsDomainCapabilities } elseif($IsUpdate) { $params.TlsDomainCapabilities = $null }

  # Remove parameters the cmdlet of the current Exchange version does not support
  $availableParameters = (Get-Command -Name $CommandName).Parameters.Keys

  foreach($key in @($params.Keys)) {
    if($availableParameters -notcontains $key) {
      Write-Verbose -Message ('{0} does not support parameter {1}. Skipping.' -f $CommandName, $key)
      $params.Remove($key)
    }
  }

  return $params
}

function Get-AceKey {
  # Builds a comparable key for an AD permission entry
  [CmdletBinding()]
  param(
    $Ace
  )

  return ('{0}|{1}|{2}|{3}|{4}' -f ([string]$Ace.User).ToLower(),
    (ConvertTo-Boolean -Value $Ace.Deny),
    (((ConvertTo-StringArray -Value $Ace.AccessRights) | Sort-Object) -join ','),
    (((ConvertTo-StringArray -Value $Ace.ExtendedRights) | Sort-Object) -join ','),
  (((ConvertTo-StringArray -Value $Ace.Properties) | Sort-Object) -join ','))
}

function Copy-ConnectorPermission {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)]
    $SourceConnector,
    [Parameter(Mandatory)]
    [string]$TargetIdentity
  )

  # fetch non inherited permissions from source connector
  $sourcePermissions = @(Get-ADPermission -Identity ([string]$SourceConnector.DistinguishedName) -DomainController $DomainController |
  Where-Object -FilterScript { -not (ConvertTo-Boolean -Value $_.IsInherited) })

  if($sourcePermissions.Count -eq 0) {
    Write-Verbose -Message 'Source connector has no non-inherited AD permissions'
    return
  }

  $targetRC = Get-ReceiveConnector -Identity $TargetIdentity -DomainController $DomainController -ErrorAction SilentlyContinue

  if($null -eq $targetRC) {
    # Happens with -WhatIf, when the target connector has not been created
    Write-Verbose -Message ('Target connector {0} not found. {1} AD permissions would be copied.' -f $TargetIdentity, $sourcePermissions.Count)
    return
  }

  $targetDN = [string]$targetRC.DistinguishedName

  # Fetch existing permissions on target to avoid duplicate ACE errors
  $existingAceKeys = @(Get-ADPermission -Identity $targetDN -DomainController $DomainController |
      Where-Object -FilterScript { -not (ConvertTo-Boolean -Value $_.IsInherited) } |
  ForEach-Object -Process { Get-AceKey -Ace $_ })

  Write-Verbose -Message 'Adding AD permissions'

  $added = 0
  $skipped = 0

  foreach($ace in $sourcePermissions) {

    $user = [string]$ace.User

    if($user -eq '') {
      continue
    }

    if($existingAceKeys -contains (Get-AceKey -Ace $ace)) {
      Write-Verbose -Message ('Permission for {0} already exists on target. Skipping.' -f $user)
      $skipped++
      continue
    }

    $aceParams = @{
      Identity         = $targetDN
      User             = $user
      Deny             = (ConvertTo-Boolean -Value $ace.Deny)
      DomainController = $DomainController
    }

    $accessRights = ConvertTo-StringArray -Value $ace.AccessRights
    $extendedRights = ConvertTo-StringArray -Value $ace.ExtendedRights
    $adProperties = ConvertTo-StringArray -Value $ace.Properties
    $inheritanceType = [string]$ace.InheritanceType

    if($accessRights.Count -gt 0) { $aceParams.AccessRights = $accessRights }
    if($extendedRights.Count -gt 0) { $aceParams.ExtendedRights = $extendedRights }
    if($adProperties.Count -gt 0) { $aceParams.Properties = $adProperties }
    if($inheritanceType -ne '') { $aceParams.InheritanceType = $inheritanceType }

    if($PSCmdlet.ShouldProcess($TargetIdentity, ('Add AD permission for {0} ({1})' -f $user, ($accessRights + $extendedRights -join ', ')))) {
      try {
        $null = Add-ADPermission @aceParams -ErrorAction Stop
        $added++
      }
      catch {
        Write-Warning -Message ('Failed to add AD permission for {0} on {1}: {2}' -f $user, $TargetIdentity, $_.Exception.Message)
      }
    }
  }

  Write-Host ('AD permissions on {0}: {1} added, {2} already existing' -f $TargetIdentity, $added, $skipped)
}

function Wait-ForDirectory {
  [CmdletBinding()]
  param()

  if($SecondsToWait -gt 0) {
    Write-Host ('Wait {0} seconds for domain controller to update' -f $SecondsToWait)
    Start-Sleep -Seconds $SecondsToWait
  }
}

function Copy-ToServer {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)]
    $SourceConnector,
    [Parameter(Mandatory)]
    $SourceServerObject,
    [Parameter(Mandatory)]
    [string]$TargetServerName
  )

  $targetServerObject = Get-ExchangeServer -Identity $TargetServerName -DomainController $DomainController -ErrorAction SilentlyContinue

  if($null -eq $targetServerObject) {
    Write-Warning -Message ('Target server {0} cannot be found.' -f $TargetServerName)
    return
  }

  $targetName = [string]$targetServerObject.Name

  if($targetName -eq [string]$SourceServerObject.Name) {
    Write-Warning -Message 'Source and target server are identical. Skipping.'
    return
  }

  $targetIdentity = '{0}\{1}' -f $targetName, $ConnectorName

  Write-Host ('Working on [{0}] and receive connector [{1}]' -f $targetName.ToUpper(), $ConnectorName)

  # Replace the source server FQDN with the target server FQDN
  $targetFqdn = [string]$SourceConnector.Fqdn
  $sourceServerFqdn = [string]$SourceServerObject.Fqdn

  if(($targetFqdn -eq $sourceServerFqdn) -or ($targetFqdn -like ('{0}.*' -f $SourceServerObject.Name))) {
    Write-Verbose -Message 'Source connector uses server Fqdn. Changing to target server Fqdn.'
    $targetFqdn = ([string]$targetServerObject.Fqdn).ToLower()
  }

  $targetRC = Get-ReceiveConnector -Identity $targetIdentity -DomainController $DomainController -ErrorAction SilentlyContinue

  if($null -eq $targetRC) {

    # create new receive connector
    $params = Get-ReceiveConnectorParameter -Connector $SourceConnector -TargetFqdn $targetFqdn -CommandName 'New-ReceiveConnector'

    $params.Name = [string]$SourceConnector.Name
    $params.Server = $targetName
    $params.DomainController = $DomainController

    if($MoveToFrontend) {
      # Move receive connector to FrontEnd Transport
      $params.TransportRole = 'FrontendTransport'
    }
    else {
      $params.TransportRole = [string]$SourceConnector.TransportRole
    }

    if($ResetBindings) {
      # Reset network bindings to listen on all adapters using port 25
      $params.Bindings = @('0.0.0.0:25')
    }
    else {
      $params.Bindings = ConvertTo-StringArray -Value $SourceConnector.Bindings

      $serverSpecificBindings = @($params.Bindings | Where-Object -FilterScript { Test-ServerSpecificBinding -Binding $_ })

      if($serverSpecificBindings.Count -gt 0) {
        Write-Warning -Message ('Source connector uses server specific bindings ({0}). Use -ResetBindings or adjust the bindings on {1} afterwards.' -f ($serverSpecificBindings -join ', '), $targetName)
      }
    }

    # ExternalAuthoritative requires ExchangeServers permission group during creation
    $restorePermissionGroups = $null

    if((($params.AuthMechanism -split ',').Trim() -contains 'ExternalAuthoritative') -and ($params.PermissionGroups -ne 'ExchangeServers')) {
      Write-Verbose -Message 'Connector is set to ExternalAuthoritative and needs special treatment.'
      $restorePermissionGroups = $params.PermissionGroups
      $params.PermissionGroups = 'ExchangeServers'
    }

    if($PSCmdlet.ShouldProcess($targetIdentity, 'Create receive connector')) {
      try {
        $null = New-ReceiveConnector @params -ErrorAction Stop
      }
      catch {
        Write-Warning -Message ('Failed to create receive connector {0}: {1}' -f $targetIdentity, $_.Exception.Message)
        return
      }

      if($null -ne $restorePermissionGroups) {
        Wait-ForDirectory
        Write-Verbose -Message ('Updating PermissionGroups for {0}' -f $targetIdentity)
        Set-ReceiveConnector -Identity $targetIdentity -PermissionGroups $restorePermissionGroups -DomainController $DomainController
      }

      if($CopyPermissions) {
        Wait-ForDirectory
        Copy-ConnectorPermission -SourceConnector $SourceConnector -TargetIdentity $targetIdentity
      }
    }
  }
  else {
    # target connector already exists
    Write-Host 'Target connector already exists.'

    if(($UpdateExistingConnector) -or ((Request-Choice -Caption ('Do you want to UPDATE the receive connector {0} on server {1}?' -f $ConnectorName, $targetName)) -eq 0)) {

      $params = Get-ReceiveConnectorParameter -Connector $SourceConnector -TargetFqdn $targetFqdn -CommandName 'Set-ReceiveConnector' -IsUpdate

      $params.Identity = $targetIdentity
      $params.DomainController = $DomainController

      if($PSCmdlet.ShouldProcess($targetIdentity, 'Update receive connector')) {

        Write-Host ('Updating connector on server {0}' -f $targetName)

        try {
          Set-ReceiveConnector @params -ErrorAction Stop
        }
        catch {
          Write-Warning -Message ('Failed to update receive connector {0}: {1}' -f $targetIdentity, $_.Exception.Message)
          return
        }
      }

      if($CopyPermissions) {
        Copy-ConnectorPermission -SourceConnector $SourceConnector -TargetIdentity $targetIdentity
      }
    }
  }
}

function Copy-ToAllServer {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)]
    $SourceConnector,
    [Parameter(Mandatory)]
    $SourceServerObject
  )

  Write-Verbose -Message 'Copy receive connector to all other Exchange servers'

  # All Exchange 2013+ servers hosting the Mailbox role (Exchange 2016, 2019, SE), Edge Transport servers excluded
  $targetServers = @(Get-ExchangeServer -DomainController $DomainController |
      Where-Object -FilterScript { ([string]$_.AdminDisplayVersion -match '^Version 15\.') -and ([string]$_.ServerRole -match 'Mailbox') -and ([string]$_.Name -ne [string]$SourceServerObject.Name) } |
  Sort-Object -Property Name)

  foreach($server in $targetServers) {
    Copy-ToServer -SourceConnector $SourceConnector -SourceServerObject $SourceServerObject -TargetServerName ([string]$server.Name)
  }

  Write-Verbose -Message 'Finished copying connector to all modern Exchange servers'
}

### MAIN ----------------------------------

if($null -eq (Get-Command -Name Get-ReceiveConnector -ErrorAction SilentlyContinue)) {
  Write-Warning -Message 'Exchange cmdlets are not available. Please run this script in the Exchange Management Shell.'
  return
}

if((-not $CopyToAllOther) -and ($TargetServer -eq '')) {
  Write-Warning -Message 'You need to either specify a dedicated target server using the -TargetServer parameter or select the -CopyToAllOther switch.'
  return
}

if($ViewEntireForest) {
  Write-Verbose -Message ('Setting ADServerSettings -ViewEntireForest {0}' -f $true)
  Set-ADServerSettings -ViewEntireForest $true
}

$sourceServerObject = Get-ExchangeServer -Identity $SourceServer -DomainController $DomainController -ErrorAction SilentlyContinue

if($null -eq $sourceServerObject) {
  Write-Warning -Message ('Source server {0} cannot be found.' -f $SourceServer)
  return
}

# Read the source connector once
$sourceRC = Get-ReceiveConnector -Identity ('{0}\{1}' -f $sourceServerObject.Name, $ConnectorName) -DomainController $DomainController -ErrorAction SilentlyContinue

if($null -eq $sourceRC) {
  Write-Warning -Message ('Source connector {0}\{1} cannot be accessed or does not exist!' -f $SourceServer, $ConnectorName)
  return
}

if($TargetServer -ne '') {
  # Copy to a single Exchange server
  Copy-ToServer -SourceConnector $sourceRC -SourceServerObject $sourceServerObject -TargetServerName $TargetServer
}
else {
  # Copy to all Exchange 2016/2019/SE servers
  Copy-ToAllServer -SourceConnector $sourceRC -SourceServerObject $sourceServerObject
}