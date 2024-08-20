# Test-DomainAvailability.ps1

Check the availability of a domain in a selected Office 365 region.

## Description

The script queries the login uri for the selected Office 365 region. The response contains metadata about the domain queried.

If the domain already exists in the specified region the metadata contains information if the domain is verified and/or federated

Load function into your current PowerShell session:

``` PowerShell
. .\Test-DomainAvailability.ps1
```

## Parameters

### DomainName

The domain name you want to verify. Example: example.com

### LookupRegion

The Office 365 region where you want to verify the domain.
Currently implemented: Global, Germany, China

## Examples

``` PowerShell
Test-DomainAvailability -DomainName example.com 
```

Test domain availability in the default region - Office 365 Global

``` PowerShell
Test-DomainAvailability -DomainName example.com -LookupRegion China
```

Test domain availability in Office 365 China

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)