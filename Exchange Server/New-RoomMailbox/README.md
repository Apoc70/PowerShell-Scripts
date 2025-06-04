# New-RoomMailbox.ps1

Creates a new room mailbox, security groups for full access and send-as permission and adds the security groups to the room mailbox configuration.

## Description

This scripts creates a new room mailbox and additonal security groups for full access and and send-as delegation. The security groups are created using a confgurable naming convention.
You can add existing groups to the full access and send-as to a room mailbox. This is useful, if you have a room management department and want to grant permissions.
The script adds the room mailbox to an existing room list, if configured.

## Parameters

### RoomMailboxName

Name attribute of the new room mailbox

### RoomMailboxDisplayName

Display name attribute of the new room mailbox

### RoomMailboxAlias

Alias attribute of the new room mailbox

### RoomMailboxSmtpAddress

Primary SMTP address attribute the new room mailbox

### DepartmentPrefix

Department prefix for automatically generated security groups (optional)

### GroupFullAccessMembers

String array containing full access members

### GroupSendAsMembers

String array containing send as members

### GroupCalendarBookingMembers

String array containing users having calendar booking rights

### RoomPhoneNumber

Phone number of a phone located in the room, this value will show in the Outlook room list

### RoomList

Add the new room mailbox to this existing room list

### AutoAccept

Set room mailbox to automatically accept booking requests

### Language

Locale setting for calendar regional configuration language, e.g., de-DE, en-US

## Examples

``` PowerShell
.\New-RoomMailbox.ps1 -RoomMailboxName "MB - Conference Room" -RoomMailboxDisplayName "Board Conference Room" -RoomMailboxAlias "MB-ConferenceRoom" -RoomMailboxSmtpAddress "ConferenceRoom@mcsmemail.de" -DepartmentPrefix "C"
```
Create a new room mailbox, empty full access and empty send-as security groups

## Credits

Written by: Thomas Stensitzki

### Stay connected

* Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
* Bluesky: [https://bsky.app/profile/stensitzki.eu](https://bsky.app/profile/stensitzki.eu)
* LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
* YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
* LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Microsoft 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Website: [https://granikos.eu](https://granikos.eu)
* Bluesky: [https://bsky.app/profile/granikos.eu](https://bsky.app/profile/granikos.eu)
