# Move-MigrationUser.ps1

 This script creates a new migration batch and moves migration users from one batch to the new batch.

## Description

This script moves Exchange Online migration users to a new migration batch. The script creates the new batch automatically. The name of the new batch is based on the following:

- Custom prefix, provided by the BatchName parameter
- Batch completion date
- Source batch name

The new batch name helps you to easily identify the planned completion date and source of the migration users. This is helpful during an agile mailbox migration process.

## Requirements

- Exchange Online Management Shell v2

## Parameters

### Users

List of migration user email addresses that should move to a new migration batch

### UsersCsvFile

Path to a CSV file containing a migration users, one user migration email address per line

### BatchName

The name of the new migration batch. The BatchName is the first part of the final full Name, e.g. BATCHNAME_2022-08-17_SOURCEBATCHNAME

### Autostart

Switch indicating that the new migration batch should start automatically

### AutoComplete

Switch indicating that the new migration should complete migration automatically
NOT implemented yet

### CompleteDateTime

[DateTime] defining the completion date and time for the new batch

### NotificationEmails

Email addresses for batch notification emails

### DateTimePattern

The string pattern used for date information in the batch name

## Example

``` PowerShell
./Move-MigrationUsers -Users JohnDoe@varunagroup.de,JaneDoe@varunagroup.de -CompleteDateTime '2022/08/31 18:00'
```

Move two migration users from the their current migration  batch to a new migration batch

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

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
