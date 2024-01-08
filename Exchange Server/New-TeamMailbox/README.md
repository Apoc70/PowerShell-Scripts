# New-TeamMailbox.ps1

Creates a new shared mailbox, security groups for full access and send-as permission and adds the security groups to the shared mailbox configuration.

## Description

This scripts creates a new shared mailbox (aka team mailbox) and security groups for full access and and send-as delegation. Security groups are created using a naming convention.

## Parameters

### TeamMailboxName

Name attribute of the new team mailbox

### TeamMailboxDisplayName

Display name attribute of the new team mailbox

### TeamMailboxAlias

Alias attribute of the new team mailbox

### TeamMailboxSmtpAddress

Primary SMTP address attribute the new team mailbox

### DepartmentPrefix

Department prefix for automatically generated security groups (optional)

### GroupFullAccessMembers

String array containing full access members

### GroupFullAccessMembers

String array containing send as members

## Examples

``` PowerShell
.\New-TeamMailbox.ps1 -TeamMailboxName "TM-Exchange Admins" -TeamMailboxDisplayName "Exchange Admins" -TeamMailboxAlias "TM-ExchangeAdmins" -TeamMailboxSmtpAddress "ExchangeAdmins@mcsmemail.de" -DepartmentPrefix "IT"
```

Create a new team mailbox, empty full access and empty send-as security groups

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