<#
    .SYNOPSIS
    Configure Exchange Server 2016+ Virtual Directory Url Settings

    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 2.2, 2026-02-04

    Use GitHub for comments, suggestions, and issues.

    .LINK
    https://scripts.granikos.eu

    .DESCRIPTION
    Exchange Server virtual directories (vDirs) require a proper configuration of
    internal and external Urls. This is even more important in a co-existence
    scenario with legacy Exchange Server versions.

    .NOTES
    Requirements
    - Windows Server 2016+
    - Exchange Management Shell for Exchange Server 2016+

    Revision History
    --------------------------------------------------------------------------------
    1.0 Initial community release
    2.0 Updated for Exchange Server 2016, 2019, SE
    2.1 PowerShell Hygiene
    2.2 Added option to configure AutoDiscover Virtual Directory Urls

    .PARAMETER InternalUrl
    The internal url FQDN with protocol definition, ie. https://mobile.mcsmemail.de

    .PARAMETER ExternalUrl
    The internal url FQDN with protocol definition, ie. https://mobile.mcsmemail.de 

    .PARAMETER AutoDiscoverUrl
    The internal url FQDN for Autodiscover with protocol definition, ie. https://

    .PARAMETER ServerName
    The Exchange Server name to configure. If not specified, all Exchange 2016+

    .EXAMPLE
    Configure internal and external url for different host headers
    .\Set-VirtualDirectoryUrl.ps1 -InternalUrl https://internal.mcsmemail.de -ExternalUrl https://mobile.mcsmemail.de

    .EXAMPLE
    Configure AutoDiscover url only
    .\Set-VirtualDirectoryUrl.ps1 -AutoDiscoverUrl https://autodiscover.mcsmemail.de

#>
[CmdletBinding()]
Param(
  [string]$InternalUrl = '',
  [string]$ExternalUrl = '',
  [string]$AutoDiscoverUrl = '',
  [string]$ServerName = ''  
)

# Helper function to trim trailing slash
function Remove-TrailingSlash {
  param([string]$Url)
  return $Url.TrimEnd('/')
}

if($ServerName -ne '') {
  # Fetch Exchange Server
  $exchangeServers = Get-ExchangeServer -Identity $ServerName
}
else {
  # Fetch Exchange Server 2016+ Servers
  $exchangeServers = Get-ExchangeServer | Where-Object {($_.IsE15OrLater -eq $true) -and ($_.ServerRole -ilike '*Mailbox*')}
}

if($AutoDiscoverUrl -ne '') {
  
  # Trim trailing "/"
  $AutoDiscoverUrl = Remove-TrailingSlash -Url $AutoDiscoverUrl
  
  # Set AutoDiscover Url for Exchange CAS SCP
  Write-Output 'Changing AutodiscoverServiceInternalUri settings'
  Write-Output ('New AutodiscoverServiceInternalUri: {0}/autodiscover/autodiscover.xml' -f $AutoDiscoverUrl)
  $null = $exchangeServers | ForEach-Object{ Set-ClientAccessService -Identity $_.Name -AutodiscoverServiceInternalUri ('{0}/autodiscover/autodiscover.xml' -f $AutoDiscoverUrl) -Confirm:$false}

}
else {
  Write-Output 'No AutoDiscoverUrl specified. Skipping AutoDiscover configuration.'
}

if($InternalUrl -ne '' -and $null -ne $exchangeServers) {

  # Trim trailing "/"
  $InternalUrl = Remove-TrailingSlash -Url $InternalUrl

  Write-Output 'The script configures the following servers:'
  $exchangeServers | Format-Table -Property Name, AdminDisplayVersion -AutoSize

  Write-Output 'Changing InternalUrl settings'
  Write-Output ('New INTERNAL Url: {0}' -f $InternalUrl)

  # Set Internal Urls
  $null = $exchangeServers | Get-WebServicesVirtualDirectory -ADPropertiesOnly | Set-WebServicesVirtualDirectory -InternalUrl ('{0}/ews/exchange.asmx' -f $InternalUrl) -Confirm:$false
  $null = $exchangeServers | Get-OabVirtualDirectory -ADPropertiesOnly | Set-OabVirtualDirectory -InternalUrl ('{0}/oab' -f $InternalUrl) -Confirm:$false
  $null = $exchangeServers | Get-OwaVirtualDirectory -ADPropertiesOnly | Set-OwaVirtualDirectory -InternalUrl ('{0}/owa' -f $InternalUrl) -Confirm:$false
  $null = $exchangeServers | Get-EcpVirtualDirectory -ADPropertiesOnly | Set-EcpVirtualDirectory -InternalUrl ('{0}/ecp' -f $InternalUrl) -Confirm:$false
  $null = $exchangeServers | Get-MapiVirtualDirectory -ADPropertiesOnly | Set-MapiVirtualDirectory -InternalUrl ('{0}/mapi' -f $InternalUrl) -Confirm:$false
  $null = $exchangeServers | Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | Set-ActiveSyncVirtualDirectory -InternalUrl ('{0}/Microsoft-Server-ActiveSync' -f $InternalUrl) -Confirm:$false
  Write-Output 'InternalUrl changed'
}

if($ExternalUrl -ne '' -and $null -ne $exchangeServers) {

  # Trim trailing "/"
  $ExternalUrl = Remove-TrailingSlash -Url $ExternalUrl

  Write-Output 'Changing ExternalUrl settings'
  Write-Output ('New EXTERNAL Url: {0}' -f $ExternalUrl)

  # Set External Urls
  $null = $exchangeServers | Get-WebServicesVirtualDirectory -ADPropertiesOnly| Set-WebServicesVirtualDirectory -ExternalUrl ('{0}/ews/exchange.asmx' -f $ExternalUrl) -Confirm:$false
  $null = $exchangeServers | Get-OabVirtualDirectory -ADPropertiesOnly | Set-OabVirtualDirectory -ExternalUrl ('{0}/oab' -f $ExternalUrl) -Confirm:$false
  $null = $exchangeServers | Get-OwaVirtualDirectory -ADPropertiesOnly | Set-OwaVirtualDirectory -ExternalUrl ('{0}/owa' -f $ExternalUrl) -Confirm:$false
  $null = $exchangeServers | Get-EcpVirtualDirectory -ADPropertiesOnly | Set-EcpVirtualDirectory -ExternalUrl ('{0}/ecp' -f $ExternalUrl) -Confirm:$false
  $null = $exchangeServers | Get-MapiVirtualDirectory -ADPropertiesOnly | Set-MapiVirtualDirectory -ExternalUrl ('{0}/ecp' -f $ExternalUrl) -Confirm:$false
  $null = $exchangeServers | Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | Set-ActiveSyncVirtualDirectory -ExternalUrl ('{0}/Microsoft-Server-ActiveSync' -f $ExternalUrl) -Confirm:$false
}

if(($InternalUrl -ne '') -or ($ExternalUrl -ne '')) {
  # Query Settings
  Write-Output ''
  Write-Output 'Current Url settings for CAS AutodiscoverServiceInternalUri'
  $exchangeServers | Get-ClientAccessService | Format-List -Property Server, Identity, AutodiscoverServiceInternalUri
  Write-Output 'Current Url settings for Web Services Virtual Directory'
  $exchangeServers | Get-WebServicesVirtualDirectory -ADPropertiesOnly | Format-List -Property Server, Identity, InternalUrl, ExternalUrl
  Write-Output 'Current Url settings for OAB Virtual Directory'
  $exchangeServers | Get-OabVirtualDirectory -ADPropertiesOnly | Format-List -Property Server, Identity, InternalUrl, ExternalUrl
  Write-Output 'Current Url settings for OWA Virtual Directory'
  $exchangeServers | Get-OwaVirtualDirectory -ADPropertiesOnly | Format-List -Property Server, Identity, InternalUrl,ExternalUrl
  Write-Output 'Current Url settings for ECP Virtual Directory'
  $exchangeServers | Get-EcpVirtualDirectory -ADPropertiesOnly | Format-List -Property Server, Identity, InternalUrl,ExternalUrl
  Write-Output 'Current Url settings for MAPI Virtual Directory'
  $exchangeServers | Get-MapiVirtualDirectory -ADPropertiesOnly | Format-List -Property Server, Identity, InternalUrl,ExternalUrl
  Write-Output 'Current Url settings for ActiveSync Virtual Directory'
  $exchangeServers | Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | Format-List -Property Server, Identity, internalurl, ExternalUrl
}
else {
  Write-Output 'Nothing changed!'
}
