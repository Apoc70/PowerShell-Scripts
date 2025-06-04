# Add-CustomCalendarItems.ps1

Send a single monitoring email to a recipient

## Description

This script send a single monitoring email to a recipient. The script is intended to be used as a monitoring script.
The message body contains a GUID and the current date and time in ticks.

Use the Add-EntraIdAppRegistration.ps1 script to create a custom application registration in Entra ID.

Adjust the settings in the Settings.xml file to match your environment.

When using application permissions for Microsoft Grapg, consider restricting access to the application to specific users or groups:
[https://bit.ly/LimitExoAppAccess](https://bit.ly/LimitExoAppAccess)

## Requirements

- PowerShell 7.1+
- GlobalFunctions PowerShell module
- Registered Entra ID application with access to Microsoft Graph

## Parameters

### SettingsFileName

The file name of the settings file located in the script directory.

## Example

``` PowerShell
.\Add-CustomCalendarItems.ps1 -SettingsFileName CustomSettings.xml
```

Send a monitoring email using the selected settings file

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

THe supporting script Add-EntraIdAppRegistration is based on content published by [Andres Bohren](https://blog.icewolf.ch/archive/2022/12/02/create-azure-ad-app-registration-with-microsoft-graph-powershell)

### Stay connected

* Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
* Bluesky: [https://bsky.app/profile/stensitzki.eu](https://bsky.app/profile/stensitzki.eu)
* LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
* YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
* LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Microsoft 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Website: [https://granikos.eu](https://granikos.eu)
* Bluesky: [https://bsky.app/profile/granikos.eu](https://bsky.app/profile/granikos.eu)
