# New-TeamMailbox.ps1

Creates a new shared mailbox, security groups for full access and send-as permission and adds the security groups to the shared mailbox configuration.

## Description

This scripts creates a new shared mailbox (aka team mailbox) and security groups for full access and and send-as delegation. Security groups are created using a naming convention.

Starting with v1.4 the script sets the sAMAccountName of the security groups to the group name to avoid numbered name extension of sAMAccountName.

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

### Stay connected

- My Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
- Bluesky: [https://bsky.app/profile/stensitzki.bsky.social](https://bsky.app/profile/stensitzki.bsky.social)
- LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
- YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
- LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Website: [https://www.granikos.eu](https://www.granikos.eu)
- Bluesky: [https://bsky.app/profile/granikos.bsky.social](https://bsky.app/profile/granikos.bsky.social)