# Get-SoftwareInventory.ps1

Script to generate an html reports of installed software, installed updates and installed components on a remote computer

## Description

This script utilizes the slightly modified function Get-RemoteProgram by Jaap Brasser to fetch information about installed software.

## Parameters

### RemoteComputer

Name of the computer to fetch data from (default: local computer)

### SoftwareFilter

Filter for installed software. Only software matching the filter will be included in the report.

### ProgramsOnly

Switch to only fetch installed programs and exclude updates and components

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
.\Get-SoftwareInventory.ps1 -SendMail -MailFrom postmaster@mcsmemail.de -MailTo it-support@mcsmemail.de -MailServer mymailserver.mcsmemail.de -RemoteComputer SERVER01,SERVER02
```

Get software information about SERVER01, SERVER02 and send email report

``` PowerShell
.\Get-SoftwareInventory.ps1 -RemoteComputer SERVER01,SERVER02
```

Get software information about SERVER01, SERVER02 and write report to disk only

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

## Additional Credits

- Jaap Brasser
- Georges Zwingelstein
