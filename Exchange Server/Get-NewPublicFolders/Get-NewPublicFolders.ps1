<# 
  .SYNOPSIS 

  Get all public folders created during the las X days

  Thomas Stensitzki 

  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

  Version 2.0, 2024-10-15

  Please send ideas, comments and suggestions to support@granikos.eu 

  .LINK 
  https://scripts.granikos.eu

  .DESCRIPTION 

  This script gathers all public folders created during the last X days and exportes the gathered data to a CSV file.

  .NOTES 

  Requirements Legacy Public Folder
  - Windows Server 2008R2+ 
  - Exchange 2010/2013 Management Shell (aka EMS)

    Requirements Modern Public Folder
  - Windows Server 2012R2+
  - Exchange 2013+ Management Shell (aka EMS)

  Revision History 
  -------------------------------------------------------------------------------- 
  1.0 Initial community release
  2.0 Enhanced reporting for modern public folders
   
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
  [int]$Days = 30,
  [Parameter(ParameterSetName='Legacy')]
  [switch]$Legacy,
  [Parameter(ParameterSetName='Legacy')]
  [string]$ServerName = 'MYSERVER',
  [switch]$FetchFolderPermissions,
  [switch]$SendMail,
  [string]$MailFrom = "",
  [string]$MailTo = "",
  [string]$MailServer = ""
)

# Fetch new public folders created over the last 7 days
$CreationDate = (Get-Date).AddDays(-($Days))

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

# Use current date as file timestamp
$now = Get-Date -Format 'yyyy-MM-dd'
$CsvFilePath = Join-Path -Path $ScriptDir -ChildPath ('Get-NewPublicFolders {0}.csv' -f ($now)) 

$cssFile = Join-Path -Path $ScriptDir -ChildPath styles.css
$now = Get-Date -Format F
$reportTitle = "New Public Folder Report - $($now)"

# Gather legacy public folder statistics
if($Legacy) { 
  # Query lagacy public folders

  Get-PublicFolderStatistics -Server $ServerName | Where-Object{$_.CreationTime -ge $CreationDate} | Select-Object -Property FolderPath,Name,ItemCount | Sort-Object -Property FolderPath | Export-Csv -Path $CsvFilePath -Encoding UTF8 -NoTypeInformation -Force -Delimiter '|'
}
else {
  # Query modern public folders

  $PublicFolder = Get-PublicFolderStatistics -ResultSize Unlimited | Where-Object{$_.CreationTime -ge $CreationDate} | Select-Object -Property FolderPath,Name,ItemCount,CreationTime,LastModificationTime | Sort-Object -Property FolderPath 

  $exportFolders = New-Object System.Collections.ArrayList -ArgumentList ( ($PublicFolder | Measure-Object).Count )

  if($FetchFolderPermissions) {

    Write-Verbose 'Fetching Public Folder Permissions'

    foreach($Folder in $PublicFolder) {

        Write-Verbose ('Processiong {0}' -f $Folder.Name)

        $folderPermissions = Get-PublicFolderClientPermission "\Test" | Where-Object{$_.AccessRights -like 'Owner'} | Select-Object User

        $property = [ordered]@{
            FolderPath  = $Folder.FolderPath
            Name        = $Folder.Name
            ItemCount   = $Folder.ItemCount      
            CreationTime = $Folder.CreationTime
            LastModificationTime = $Folder.LastModificationTime
            Owner       = [string]::Join(', ',$a.User.DisplayName)
        }

        $folderObject = New-Object -TypeName PSObject -Property $property

        $null = $exportFolders.Add($folderObject)
    }

  }
  else {

    foreach($Folder in $PublicFolder) {

        $property = [ordered]@{
            FolderPath  = $Folder.FolderPath
            Name        = $Folder.Name
            ItemCount   = $Folder.ItemCount
            CreationTime = $Folder.CreationTime
            LastModificationTime = $Folder.LastModificationTime
        }

        $folderObject = New-Object -TypeName PSObject -Property $property

        $null = $exportFolders.Add($folderObject)
    }
  }
}

# Export to CSV
if(($PublicFolder | Measure-Object).Count -gt 0) {

  $exportFolders | Export-Csv -Path $CsvFilePath -Encoding UTF8 -NoTypeInformation -Force -Delimiter '|'

  $PublicFolderCount = ($PublicFolder | Measure-Object).Count
}
else {
  $PublicFolderCount = 0

  Write-Verbose -Message 'Nothing to export.'
}

if($SendMail) {

  if( $PublicFolderCount -gt 0) {
    # new public fodlers found
    $html = $exportFolders | ConvertTo-Html -Fragment -PreContent "<h2>New Public Folders</h2>"
  }
  else {
    # no new public folders found
    $html = '<p>No new public folders found.</p>'
  }

$head = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($reportTitle)</title>
<style type="text/css">$(Get-Content $cssFile)</style>
<body><h1 align=""center"">$($reportTitle)</h1>
<p>Public folder timeframe: <strong>$($Days) days</strong></p>
<p>New public folders created: <strong>$($PublicFolderCount)</strong></p>
"@

[string]$htmlreport = ConvertTo-Html -Body $html -Head $head -Title $reportTitle

Send-Mail -From $MailFrom -To $MailTo -SmtpServer $MailServer -MessageBody $htmlreport -Subject $reportTitle
}