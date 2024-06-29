# Add-CustomCalendarItems.ps1

Add calendar items to the default calendar of users in a security group

## Description

The script reads a JSON file with event data and creates calendar items in the default calendar of users in a security group.

Use the Add-EntraIdAppRegistration.ps1 script to create a custom application registration in Entra ID.

Adjust the settings in the Settings.xml file to match your environment.

See corresponging blog post [LINK TBD]()

## Requirements

- PowerShell 7.1+
- GlobalFunctions PowerShell module
- Registered Entra ID application with access to Microsoft Graph

## Parameters

### EventFileName

The name of the JSON file containing the event data located in the script directory.

### SettingsFileName

The file name of the settings file located in the script directory.

## Example

``` PowerShell
.\Add-CustomCalendarItems.ps1 -EventFileName CustomEvents.json SettingsFileName CustomSettings.xml
```

Create calendar items for users based on the JSON file CustomEvents.json and the settings file CustomSettings.xm

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

THe supporting script Add-EntraIdAppRegistration is based on content published by [Andres Bohren](https://blog.icewolf.ch/archive/2022/12/02/create-azure-ad-app-registration-with-microsoft-graph-powershell)

### Stay connected

- My Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
- Bluesky: [https://bsky.app/profile/stensitzki.bsky.social](https://bsky.app/profile/stensitzki.bsky.social)
- LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
- YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
- LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Website: [https://granikos.eu](https://www.granikos)
- Bluesky: [https://bsky.app/profile/granikos.bsky.social](https://bsky.app/profile/granikos.bsky.social)
