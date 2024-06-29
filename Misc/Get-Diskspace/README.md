# Get-Diskspace.ps1

Fetches disk/volume information from a given computer

## Description

This script fetches disk/volume information from a given computer and displays

- Volume name
- Capacity
- Free Space
- Boot Volume Status
- System Volume Status
- File System Type

With -SendMail switch no data is returned to the console.

## Requirements

- Windows Server 2012R2, 2016, 2019
- Exchange Server 2013+ (for AllExchangeServer switch)
- WMI access to remote computers

## Parameters

### ComputerName

Can of the computer to fetch disk information from

### Unit

Target unit for disk space value (default = GB)

### AllExchangeServer

Switch to fetch disk space data from all Exchange Servers

### SendMail

Switch to send an Html report

### MailFrom

Email address of report sender

### MailTo

Email address of report recipient

### MailServer

SMTP Server for email report

## Examples

``` PowerShell
.\Get-Diskpace.ps1 -ComputerName MYSERVER
```

Get disk information from computer MYSERVER

``` PowerShell
.\Get-Diskpace.ps1 -ComputerName MYSERVER -Unit MB
```

Get disk information from computer MYSERVER in MB

``` PowerShell
.\Get-Diskpace.ps1 -AllExchangeServer -SendMail -MailFrom postmaster@sedna-inc.com -MailTo exchangeadmin@sedna-inc.com -MailServer mail.sedna-inc.com
```

Get disk information from all Exchange servers and send html email

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