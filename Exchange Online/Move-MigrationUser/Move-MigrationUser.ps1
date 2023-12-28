<#
  .SYNOPSIS
  This script creates a new migration batch and moves migration users from one batch to the new batch.

  Thomas Stensitzki

  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

  Version 1.0, 2022-08-29

  Please use GitHub repository for ideas, comments, and suggestions.

  .LINK
  https://github.com/Apoc70/PowerShell-Scripts
  https://scripts.granikos.eu

  .DESCRIPTION

  This script moves Exchange Online migration users to a new migration batch. The script creates
  the new batch automatically. The name of the new batch is based on the following:
  * Custom prefix, provided by the BatchName parameter
  * Batch completion date
  * Source batch name

  The new batch name helps you to easily identify the planned completion date and source of the
  migration users. This is helpful during an agile mailbox migration process.

  .NOTES
  Requirements
  - Exchange Online Management Shell v2

  Revision History
  --------------------------------------------------------------------------------
  1.0      Initial community release

  .PARAMETER Users

  List of migration user email addresses that should move to a new migration batch

  .PARAMETER UsersCsvFile

  Path to a CSV file containing a migration users, one user migration email address per line

  .PARAMETER BatchName

  The name of the new migration batch. The BatchName is the first part of the final full Name, e.g. BATCHNAME_2022-08-17_SOURCEBATCHNAME

  .PARAMETER Autostart

  Switch indicating that the new migration batch should start automatically

  .PARAMETER AutoComplete

  Switch indicating that the new migration should complete migration automatically
  NOT implemented yet

  .PARAMETER CompleteDateTime

  [DateTime] defining the completion date and time for the new batch

  .PARAMETER NotificationEmails

  Email addresses for batch notification emails

  .PARAMETER DateTimePattern

  The string pattern used for date information in the batch name

  .EXAMPLE

  ./Move-MigrationUsers -Users JohnDoe@varunagroup.de,JaneDoe@varunagroup.de -CompleteDateTime '2022/08/31 18:00'

  Move two migration users from the their current migration  batch to a new migration batch
#>

[CmdletBinding()]
param(

    $Users = @( 'JohnDoe@varunagroup.de' ),
    # $UsersCsvFile = '',
    [string]$BatchName = 'BATCH',
    [switch]$Autostart,
    # [switch]$AutoComplete,
    [Parameter(Mandatory=$true)]
    [datetime]$CompleteDateTime,
    $NotificationEmails = 'VarunaAdmin@varunagroup.de',
    [string]$DateTimePattern = 'yyyy-MM-dd'
)

if (($Users | Measure-Object).Count -ne 0) {

    # Create name for en migration batch
    $BatchNamePrefix = ('{0}_{1}' -f $BatchName, ($CompleteDateTime.ToString($DateTimePattern)))
    Write-Verbose -Message (('New target batch name: {0}' -f $BatchName))

    # Fetching users
    Write-Verbose -Message ('Parsing {0} users' -f (($Users | Measure-Object).Count))

    # Fetch migration users and group by BatchId
    # ToDo: Pre-Check users for existence as migration user and a non-completed migration status
    $MigrationsUsers = $Users | ForEach-Object { Get-MigrationUser -Identity $_ | Select-Object Identity, BatchId } | Group-Object BatchId

    # Parse migration users
    $MigrationsUsers | ForEach-Object {

        Write-Output ('{1} - {0}'-f $_.Name, $_.Group.Identity)

        # assemble new batch name
        # adjust to your personal needs
        $NewBatchName = ('{0}_{1}' -f $BatchNamePrefix, $_.Name)

        # Create new migration batch
        if($Autostart) {
            $Batch = New-MigrationBatch -Name $NewBatchName -UserIds $_.Group.Identity -DisableOnCopy -AutoStart -NotificationEmails $NotificationEmails
        }
        else {
            $Batch = New-MigrationBatch -Name $NewBatchName -UserIds $_.Group.Identity -DisableOnCopy -NotificationEmails $NotificationEmails
        }

        # Output detailed batch information
        $Batch | Format-List

        # Set auto completion date and time if not null
        # Set time as Universal Time
        if($null -ne $CompleteDateTime) {

            Write-Verbose ('Setting CompleteAfter to: {0}' -f ($CompleteTime).ToUniversalTime())

            Set-MigrationBatch -Identity $NewBatchName -CompleteAfter ($CompleteTime).ToUniversalTime()
        }

    }
}
else {
    Write-Output 'The list of users is empty. Nothing to do today.'
}
