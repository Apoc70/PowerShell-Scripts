<#
  .SYNOPSIS
  Get all public folders created during the last X days

  Thomas Stensitzki

  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

  Version 1.1, 2024-10-01

  Use the github repository to report issues or to request additions.

  .LINK

  https://scripts.granikos.eu

  .DESCRIPTION

  This script gathers all public folders created during the last X days and exportes the gathered data to a CSV file.

  .NOTES

  Requirements
  - Windows Server 2016+
  - Exchange 2016 Management Shell (aka EMS)

  The legacy option requires the Exchange 2010 Management Shell to be installed on the local system.

  Revision History
  --------------------------------------------------------------------------------
  1.0 Initial community release
  1.1 Added support for CreationTime, LastModificationTime when exporting modern public folder statistics

  .PARAMETER Days

  Number of last X days to filter newly created public folders. Default: 14

  .PARAMETER Legacy

  Switch to define that you want to query legacy public folders

  .PARAMETER ServerName

  Name of Exchange server hostingl egacy public folders

  .EXAMPLE

  Query legacy public folder server MYPFSERVER01 for all public folders created during the last 31 days

  .\Get-NewPublicFolders.ps1 -Days 31 -ServerName MYPFSERVER01 -Legacy

  .EXAMPLE
  Query modern public folders for all public folders created during the last 31 days

  .\Get-NewPublicFolders.ps1 -Days 31

#>
[CmdletBinding()]
param(
  [int]$Days = 14,
  [Parameter(ParameterSetName = 'Legacy')]
  [switch]$Legacy,
  [Parameter(ParameterSetName = 'Legacy')]
  [string]$ServerName = 'MYSERVER'
)

# Fetch new public folders created over the last 7 days
$CreationDate = (Get-Date).AddDays( - ($Days))

# Use current date as file timestamp
$now = Get-Date -Format 'yyyy-MM-dd'
$CsvFilePath = Join-Path -Path $(Split-Path -Path $script:MyInvocation.MyCommand.Path) -ChildPath ('Get-NewPublicFolders {0}.csv' -f ($now))

# Gather legacy or modern public folder statistics
if ($Legacy) {
  # Query lagacy public folders
  Get-PublicFolderStatistics -Server $ServerName | Where-Object { $_.CreationTime -ge $CreationDate } | Select-Object -Property FolderPath, Name, ItemCount | Sort-Object -Property FolderPath | Export-Csv -Path $CsvFilePath -Encoding UTF8 -NoTypeInformation -Force -Delimiter '|'
}
else {
  # Query modern public folders
  Get-PublicFolderStatistics -ResultSize Unlimited | Where-Object { $_.CreationTime -ge $CreationDate } | Select-Object -Property FolderPath, Name, ItemCount, CreationTime, LastModificationTime | Sort-Object -Property FolderPath | Export-Csv -Path $CsvFilePath -Encoding UTF8 -NoTypeInformation -Force -Delimiter '|'
}