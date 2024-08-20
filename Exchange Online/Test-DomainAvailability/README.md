# Test-DomainAvailability.ps1

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

### Stay connected

- My Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
- Bluesky: [https://bsky.app/profile/stensitzki.bsky.social](https://bsky.app/profile/stensitzki.bsky.social)
- LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
- YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
- LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Website: [https://granikos.eu](https://www.granikos)
- Bluesky: [https://bsky.app/profile/granikos.bsky.social](https://bsky.app/profile/granikos.bsky.social)