function Test-DomainAvailability {
  <#
      .SYNOPSIS
      Check the availability of a domain in a selected Office 365 region.

      .DESCRIPTION
      The script queries the login uri for the selected Office 365 region. The response contains metadata about the domain queried.

      If the domain already exists in the specified region the metadata contains information if the domain is verified and/or federated

      Load function into your current PowerShell session:
      . .\Test-DomainAvailability.ps1

      .PARAMETER DomainName
      The domain name you want to verify. Example: example.com

      .PARAMETER LookupRegion
      The Office 365 region where you want to verify the domain.
      Currently implemented: Global, Germany, China

      .NOTES

      Author: ?
      (Source: https://blogs.technet.microsoft.com/tip_of_the_day/2017/02/16/cloud-tip-of-the-day-use-powershell-to-check-domain-availability-for-office-365-and-azure/)
      Enhancement: Thomas Stensitzki

      THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
      RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

        Version 1.1, 2024-08-20

      .LINK
      https://scripts.Granikos.eu

      .EXAMPLE
      Test-DomainAvailability -DomainName example.com

      Test domain availability in the default region - Office 365 Global

      .EXAMPLE
      Test-DomainAvailability -DomainName example.com -LookupRegion China

      Test domain availability in Office 365 China
  #>

  param(
    [Parameter(Mandatory=$true,HelpMessage='A domain name (e.g. example.com) is required')]
    [string]$DomainName,
    [ValidateSet('Global','Germany','China')]
    [string]$LookupRegion = 'Global'
  )

  # Define descriptions for status codes
  $descriptions = @{
    Unknown   = 'Domain does not exist in Office 365/Azure AD'
    Managed   = 'Domain is verified but not federated'
    Federated  = 'Domain is verified and federated'
  }

  # Define login Uris for Office 365 regions, extend as needed and add hashtable key to param validate set
  $Regions = @{
    Global = 'login.microsoftonline.com'
    Germany = 'login.microsoftonline.de'
    China = 'login.partner.microsoftonline.cn'
  }

  # Select Lookup uri
  $LookupUri = $Regions[$LookupRegion]

  if($LookupUri -ne '') {

    #
    $response = Invoke-WebRequest -Uri ('https://{0}/getuserrealm.srf?login=user@{1}&xml=1' -f $LookupUri, $DomainName)

    if($response -and $response.StatusCode -eq 200) {

      # Check namespace
      $namespaceType = ([xml]($response.Content)).RealmInfo.NameSpaceType

      New-Object -TypeName PSObject -Property @{

        DomainName = $DomainName

        NamespaceType = $namespaceType

        Details = $descriptions[$namespaceType]

      } | Select-Object -Property DomainName, NamespaceType, Details

    }
    else {
      # We were not ablte to connect to lookup uri. Do wen have an Internet connection?

      Write-Error -Message 'Domain could not be verified. Please check your connectivity to login.microsoftonline.com'

    }

  }

}