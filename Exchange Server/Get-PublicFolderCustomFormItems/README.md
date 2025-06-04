# Get-PublicFolderCustomFormItems.ps1

Get public folder items of a specific type and export the results to a CSV file.

## Description

This script retrieves all public folders under a specified path, counts the items of a specific type in each folder, and exports the results to a CSV file.
It is particularly useful for counting custom forms in public folders in an Exchange environment.
You must have the Exchange Management Shell or the Exchange Online PowerShell module installed to run this script.

## Parameters

### PublicFolderPath

The path to the public folder from which to retrieve items.

### ItemType

The type of item to count in the public folders. Default is 'IPM.Post.FORMNAME'.
Adjust this parameter to specify the item type you are interested in.
You can find the form names by using the `Get-PublicFolderItemStatistics` cmdlet.

## Examples

``` PowerShell
.\Get-PublicFolderCustomFormItems.ps1 -PublicFolderPath '\Department\HR' -ItemType 'IPM.Post.FORMNAME'
```

Retrieves all public folders under '\Department\HR', counts the items of type 'IPM.Post.FORMNAME' in each folder, and exports the results to a CSV file. Adjust the name of the item type as needed.

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

### Stay connected

* Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
* Bluesky: [https://bsky.app/profile/stensitzki.eu](https://bsky.app/profile/stensitzki.eu)
* LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
* YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
* LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Microsoft 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Website: [https://granikos.eu](https://www.granikos.eu)
* Bluesky: [https://bsky.app/profile/granikos.eu](https://bsky.app/profile/granikos.eu)
