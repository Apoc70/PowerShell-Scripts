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

## Stay connected

- Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)
- Tech & Community Podcast (DE): [http://podcast.granikos.eu](http://podcast.granikos.eu)

For more Microsoft 365, Cloud Security, and Exchange Server stuff checkout the services provided by Granikos

- Website: [https://granikos.eu](https://granikos.eu)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)