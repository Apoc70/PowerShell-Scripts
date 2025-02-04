# PowerShell Scripts

This GitHub Repository contains most of my public PowerShell scripts. In the past I used dedicated repositories per script. I will archive those repositories after moving the script to this repository. C#-related projects remain in separate repositories.

## Exchange Online

Script for Exchange Online

- [Move-MigrationUser.ps1](/Exchange%20Online/Move-MigrationUser)

  This script creates a new migration batch and moves migration users from one batch to the new batch

## Modern Exchange Server

Scripts for Exchange 2013, 2016, and 2019

- [Copy-ReceiveConnector.ps1](/Exchange%20Server/Copy-ReceiveConnector)

  Copy a selected receive connector and it's configuration and permissions to other Exchange Servers

- [Export-MessageQueue.ps1](/Exchange%20Server/Export-MessageQueue)

  Export messages from a transport queue to file system for manual replay

- [Get-ExchangeEnvironmentReport.ps1](/Exchange%20Server/Get-ExchangeEnvironmentReport)

  Creates an HTML report describing the On-Premises Exchange environment.

- [Get-RemoteSmtpServers.ps1](/Exchange%20Server/Get-RemoteSmtpServers)

  Fetch all remote SMTP servers from Exchange receive connector logs

- [Import-EdgeSubscription.ps1](/Exchange%20Server/Import-EdgeSubscription)

  Little helper script when working with Edge Transport Server subscriptions

- [New-RoomMailbox.ps1](/Exchange%20Server/New-RoomMailbox)

  This scripts creates a new room mailbox and security groups for full access and and send-as delegation. As a third security group a dedicated group for allowed users to book the new room is created. The CalenderBooking security group is only created, but not assigned to the room mailbox. Security groups are created using a naming convention.

- [New-TeamMailbox.ps1](/Exchange%20Server/New-TeamMailbox)

  Creates a new shared mailbox, security groups for full access and send-as permission and adds the security groups to the shared mailbox configuration.

- [Purge-LogFiles.ps1](/Exchange%20Server/Purge-LogFiles)

  PowerShell script for modern Exchange Server environments to clean up Exchange Server and IIS log files

- [Start-MailboxImport.ps1](/Exchange%20Server/Start-MailboxImport)

  Import one or more pst files into an exisiting mailbox or a archive


## Legacy Exchange Server

Scripts for Exchange Server 2010 and older

- TBD

## Misc

Some usefull scripts not Exchange related

- [Get-Diskspace.ps1](/Misc/Get-Diskspace)

Fetches disk/volume information from a given computer

## Network

Some useful network related scripts

- [Test-DNSRecords.ps1](/Network/Test-DNSRecords)

### Stay connected

- My Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
- Bluesky: [https://bsky.app/profile/stensitzki.eu](https://bsky.app/profile/stensitzki.eu)
- LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
- YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
- LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Website: [https://granikos.eu/](https://granikos.eu/)
- Bluesky: [https://bsky.app/profile/granikos.eu](https://bsky.app/profile/granikos.eu)