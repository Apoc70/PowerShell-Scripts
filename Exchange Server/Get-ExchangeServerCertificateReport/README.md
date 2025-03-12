# Import-EdgeSubscription.ps1

This script creates an Exchange Server SSL Certificate Report

## Description

The script queries all Exchange Servers an fetches the computer certificates.
It generates a table report of the available certificates per server.
The scripts sorts the certificates by subject name and expiry date.

## Parameters

### SendMail

Send the report as an HTML email.

### MailFrom

Sender address for result summary.

### MailTo

Recipient address for result summary.

### MailServer

SMTP Server address for sending result summary.

## Example

``` PowerShell
.\Get-ExchangeCertificateReport.ps1
```

Generate an Html report file in the script directory.

## Credits

Originally written by: Paul Cunningham

Updated by: Thomas Stensitzki

### Stay connected

- My Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
- Bluesky: [https://bsky.app/profile/stensitzki.bsky.social](https://bsky.app/profile/stensitzki.bsky.social)
- LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
- YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
- LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Website: [https://granikos.eu](https://www.granikos)
- Bluesky: [https://bsky.app/profile/granikos.bsky.social](https://bsky.app/profile/granikos.bsky.social)