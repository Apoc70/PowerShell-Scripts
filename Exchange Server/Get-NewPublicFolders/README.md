# Get-NewPublicFolders.ps1

Get all public folders created during the last X days

## Description

This script gathers all public folders created during the last X days and exportes the gathered data to a CSV file.

## Parameters

### Days

Number of last X days to filter newly created public folders. Default: 14

### Legacy

Switch to define that you want to query legacy public folders

### ServerName

Name of Exchange server hostingl egacy public folders

## Examples

``` PowerShell
.\Get-NewPublicFolders.ps1 -Days 31 -ServerName MYPFSERVER01 -Legacy
```

Query legacy public folder server MYPFSERVER01 for all public folders created during the last 31 days

``` PowerShell
.\Get-NewPublicFolders.ps1 -Days 31
```

Query modern public folders for all public folders created during the last 31 days

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

- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Bluesky: [https://bsky.app/profile/granikos.bsky.social](https://bsky.app/profile/granikos.bsky.social)