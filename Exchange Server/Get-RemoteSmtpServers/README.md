# Get-RemoteSmtpServers.ps1

Fetch all remote SMTP servers from Exchange receive connector logs

## Description

This scripts fetches remote SMTP servers by searching the Exchange receive connector logs for the EHLO string.
Fetched remote servers can be exported to a single CSV file for all receive connectors across Exchange Servers or exported to a separate CSV file per Exchange Server.

## Requirements

- Exchange Server 2010, Exchange Server 2013+

## Parameters

### Servers

List of Exchange servers, modern and legacy Exchange servers cannot be mixed

### ServersToExclude

List of host names that you want to exclude from the outout

### Backend

Search backend transport (aka hub transport) log files, instead of frontend transport, which is the default

### LegacyExchange

Search legacy Exchange servers (Exchange 2010) log file location

### ToCsv

Export search results to a single CSV file for all servers

### ToCsvPerServer

Export search results to a separate CSV file per servers

### UniqueIPs

Simplify the out list by reducing the output to unique IP address

### AddDays

File selection filter, -5 will select log files changed during the last five days. Default: -10

## Examples

``` PowerShell
.\Get-RemoteSmtpServers.ps1 -Servers SRV01,SRV02 -LegacyExchange -AddDays -4 -ToCsv
```

Search legacy Exchange servers SMTP receive log files for the last 4 days and save search results in a single CSV file

``` PowerShell
.\Get-RemoteSmtpServers.ps1 -Servers SRV03,SRV04 -AddDays -4 -ToCsv -UniqueIPs
```

Search Exchange servers SMTP receive log files for the last 4 days and save search results in a single CSV file, with unique IP addresses only


## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki | MVP

Related blog post: [https://granikos.eu/fetch-remote-smtp-servers-connecting-to-exchange/](https://granikos.eu/fetch-remote-smtp-servers-connecting-to-exchange/)

### Stay connected

- My Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
- Bluesky: [https://bsky.app/profile/stensitzki.bsky.social](https://bsky.app/profile/stensitzki.bsky.social)
- LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
- YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
- LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Bluesky: [https://bsky.app/profile/granikos.bsky.social](https://bsky.app/profile/granikos.bsky.social)