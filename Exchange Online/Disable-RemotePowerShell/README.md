# Disable-RemotePowerShell

 Disable Exchange Remote PowerShell for all users, except for members of a selected security group.

## Description

This script disables the use of RemotePoweShell und EXOPowerSHell module for Microsoft 365 users, except for members of a given security group.

## Requirements

- Exchange Online Management Shell v2

## Parameters

 ### DaysToKeep

Number of days Exchange and IIS log files should be retained, default is 30 days

### AllowRPSGroupName

Name of the Active Directory security group containing all user accounts that must have Remote PowerShell enabled.
Add all user accounts (administrators, users, service accounts) to that security group.

### LogAllowedUsers

Switch to write information about RPS allowed users to the log file.

### LogDisabledUsers

Switch to write information about users RBS disabled to the log file.

## Example

``` PowerShell
.\Disable-RemotePowerShell
```

Disable Exchange Remote PowerShell for all users using default settings, and do not write any user details to a log file.

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

### Stay connected

* Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
* Bluesky: [https://bsky.app/profile/stensitzki.eu](https://bsky.app/profile/stensitzki.eu)
* LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
* YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
* LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Microsoft 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Website: [https://granikos.eu](https://granikos.eu)
* Bluesky: [https://bsky.app/profile/granikos.eu](https://bsky.app/profile/granikos.eu)
