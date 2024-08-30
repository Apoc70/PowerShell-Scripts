<#
    .SYNOPSIS
    Configure Exchange Server 2016+ Virtual Directory Url Settings

    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 2.1, 2021-11-23

    Please send ideas, comments and suggestions to support@granikos.eu

    .LINK
    https://scripts.granikos.eu

    .DESCRIPTION
    Exchange Server virtual directories (vDirs) require a proper configuration of
    internal and external Urls. This is even more important in a co-existence
    scenario with legacy Exchange Server versions.

    .NOTES
    Requirements
    - Windows Server 2016+
    - Exchange Server 2016+

    Revision History
    --------------------------------------------------------------------------------
    1.0 Initial community release
    2.0 Updated for Exchange Server 2016, 2019, vNEXT
    2.1 PowerShell Hygiene

    .PARAMETER InternalUrl
    The internal url FQDN with protocol definition, ie. https://mobile.mcsmemail.de

    .PARAMETER ExternalUrl
    The internal url FQDN with protocol definition, ie. https://mobile.mcsmemail.de

    .EXAMPLE
    Configure internal and external url for different host headers
    .\Set-VirtualDirectoryUrl.ps1 -InternalUrl https://internal.mcsmemail.de -ExternalUrl https://mobile.mcsmemail.de

#>

Param(
  [string]$InternalUrl='',
  [string]$ExternalUrl='',
  [string]$AutodiscoverUrl='',
  [string]$ServerName = ''
)

if($ServerName -ne '') {
  # Fetch Exchange Server
  $exchangeServers = Get-ExchangeServer -Identity $ServerName
}
else {
  # Fetch Exchange Server 2016+ Servers
  $exchangeServers = Get-ExchangeServer | Where-Object {($_.IsE15OrLater -eq $true) -and ($_.ServerRole -ilike '*Mailbox*')}
}

if($InternalUrl -ne '' -and $null -ne $exchangeServers) {

  # Trim trailing "/"
  if($InternalUrl.EndsWith('/')) {
    $InternalUrl = $InternalUrl.TrimEnd('/')
  }

  Write-Output 'The script configures the following servers:'
  $exchangeServers | Format-Table -Property Name, AdminDisplayVersion -AutoSize

  Write-Output 'Changing InternalUrl settings'
  Write-Output ('New INTERNAL Url: {0}' -f $InternalUrl)

  # Set Internal Urls
  $exchangeServers | ForEach-Object{ Set-ClientAccessService -Identity $_.Name -AutodiscoverServiceInternalUri ('{0}/autodiscover/autodiscover.xml' -f $AutodiscoverUrl) -Confirm:$false}
  $exchangeServers | Get-WebServicesVirtualDirectory | Set-WebServicesVirtualDirectory -InternalUrl ('{0}/ews/exchange.asmx' -f $InternalUrl) -Confirm:$false
  $exchangeServers | Get-OabVirtualDirectory | Set-OabVirtualDirectory -InternalUrl ('{0}/oab' -f $InternalUrl) -Confirm:$false
  $exchangeServers | Get-OwaVirtualDirectory | Set-OwaVirtualDirectory -InternalUrl ('{0}/owa' -f $InternalUrl) -Confirm:$false
  $exchangeServers | Get-EcpVirtualDirectory | Set-EcpVirtualDirectory -InternalUrl ('{0}/ecp' -f $InternalUrl) -Confirm:$false
  $exchangeServers | Get-MapiVirtualDirectory | Set-MapiVirtualDirectory -InternalUrl ('{0}/mapi' -f $InternalUrl) -Confirm:$false
  $exchangeServers | Get-ActiveSyncVirtualDirectory | Set-ActiveSyncVirtualDirectory -InternalUrl ('{0}/Microsoft-Server-ActiveSync' -f $InternalUrl) -Confirm:$false

  Write-Output 'InternalUrl changed'
}

if($ExternalUrl -ne '' -and $null -ne $exchangeServers) {

  # Trim trailing "/"
  if($ExternalUrl.EndsWith('/')) {
    $ExternalUrl = $ExternalUrl.TrimEnd('/')
  }

  Write-Output 'Changing ExternalUrl settings'
  Write-Output ('New EXTERNAL Url: {0}' -f $ExternalUrl)

  # Set External Urls
  $exchangeServers | Get-WebServicesVirtualDirectory | Set-WebServicesVirtualDirectory -ExternalUrl ('{0}/ews/exchange.asmx' -f $ExternalUrl) -Confirm:$false
  $exchangeServers | Get-OabVirtualDirectory | Set-OabVirtualDirectory -ExternalUrl ('{0}/oab' -f $ExternalUrl) -Confirm:$false
  $exchangeServers | Get-OwaVirtualDirectory | Set-OwaVirtualDirectory -ExternalUrl ('{0}/owa' -f $ExternalUrl) -Confirm:$false
  $exchangeServers | Get-EcpVirtualDirectory | Set-EcpVirtualDirectory -ExternalUrl ('{0}/ecp' -f $ExternalUrl) -Confirm:$false
  $exchangeServers | Get-MapiVirtualDirectory | Set-MapiVirtualDirectory -ExternalUrl ('{0}/ecp' -f $ExternalUrl) -Confirm:$false
  $exchangeServers | Get-ActiveSyncVirtualDirectory | Set-ActiveSyncVirtualDirectory -ExternalUrl ('{0}/Microsoft-Server-ActiveSync' -f $ExternalUrl) -Confirm:$false
}

if(($InternalUrl -ne '') -or ($ExternalUrl -ne '')) {
  # Query Settings
  Write-Output ''
  Write-Output 'Current Url settings for CAS AutodiscoverServiceInternalUri'
  $exchangeServers | Get-ClientAccessServer | Format-List -Property Identity, AutodiscoverServiceInternalUri
  Write-Output 'Current Url settings for Web Services Virtual Directory'
  $exchangeServers | Get-WebServicesVirtualDirectory | Format-List -Property Identity, InternalUrl, ExternalUrl
  Write-Output 'Current Url settings for OAB Virtual Directory'
  $exchangeServers | Get-OabVirtualDirectory | Format-List -Property Identity, InternalUrl, ExternalUrl
  Write-Output 'Current Url settings for OWA Virtual Directory'
  $exchangeServers | Get-OwaVirtualDirectory | Format-List -Property Identity, InternalUrl,ExternalUrl
  Write-Output 'Current Url settings for ECP Virtual Directory'
  $exchangeServers | Get-EcpVirtualDirectory | Format-List -Property Identity, InternalUrl,ExternalUrl
  Write-Output 'Current Url settings for MAPI Virtual Directory'
  $exchangeServers | Get-MapiVirtualDirectory | Format-List -Property Identity, InternalUrl,ExternalUrl
  Write-Output 'Current Url settings for ActiveSync Virtual Directory'
  $exchangeServers | Get-ActiveSyncVirtualDirectory | Format-List -Property Identity, internalurl, ExternalUrl
}
else {
  Write-Output 'Nothing changed!'
}
