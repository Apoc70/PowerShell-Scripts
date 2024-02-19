# Import-EdgeSubscription.ps1

This script imports an Edge Subscription file for a specific Active Directory site.

## Description

The script takes two parameters: the path to the Edge Subscription file and the name of the Active Directory site.
It then imports the Edge Subscription file for the specified Active Directory site and does not create an Internet send connector.
This script is primarily used for hybrid Exchange deployments where a different internet send connector already exists.

## Parameters

### edgeSubscriptionFile

The full path to the Edge Subscription file.

### activeDirectorySite

The name of the Active Directory site for subscribign the Edge Transport Server to.


## Example

``` PowerShell
.\Import-EdgeSubscription.ps1 -edgeSubscriptionFile "C:\Import\EdgeSubscription.xml" -activeDirectorySite "Munich"
```
Import an Edge Transport subscription to an Exchange organization to Active Directory site "Munich"

## Credits

Written by: Thomas Stensitzki

### Stay connected

- My Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
- Bluesky: [https://bsky.app/profile/stensitzki.bsky.social](https://bsky.app/profile/stensitzki.bsky.social)
- LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
- YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
- LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Microsoft 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Bluesky: [https://bsky.app/profile/granikos.bsky.social](https://bsky.app/profile/granikos.bsky.social)