<#
  .SYNOPSIS
  Fetches disk/volume information from a given computer

  Thomas Stensitzki

  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

  Version 1.4, 2021-11-12

  Please use GitHub repository for ideas, comments, and suggestions.

  .LINK
  https://scripts.granikos.eu

  .DESCRIPTION
  This script fetches disk/volume information from a given computer and displays

  * Volume name
  * Capacity
  * Free Space
  * Boot Volume Status
  * System Volume Status
  * File Systemtype

  With -SendMail switch no data is returned to the console.

  .NOTES
  Requirements
  - Windows Server 2012 R2+
  - Remote WMI access
  - Exchange Server Management Shell (for AllExchangeServer switch)

  Revision History
  --------------------------------------------------------------------------------
  1.0      Initial community release
  1.1      Email reports added
  1.11     Send email issue fixed
  1.12     PowerShell hygiene applied
  1.1.3    Version number adjusted, minor PowerShell adjustments
  1.3      Tested for Windows Server 2019, minor PowerShell adjustments
  1.4      TLS 1.2 added

  .PARAMETER ComputerName
  Can of the computer to fetch disk information from

  .PARAMETER Unit
  Target unit for disk space value (default = GB)

  .PARAMETER AllExchangeServer
  Switch to fetch disk space data from all Exchange Servers

  .PARAMETER SendMail
  Switch to send an Html report

  .PARAMETER MailFrom
  Email address of report sender

  .PARAMETER MailTo
  Email address of report recipient

  .PARAMETER MailServer
  SMTP Server for email report

  .EXAMPLE
  Get disk information from computer MYSERVER

  Get-Diskpace.ps1 -ComputerName MYSERVER

  .EXAMPLE
  Get disk information from computer MYSERVER in MB

  Get-Diskpace.ps1 -ComputerName MYSERVER -Unit MB

  .EXAMPLE
  Get disk information from all Exchange servers and send html email

  Get-Diskpace.ps1 -AllExchangeServer -SendMail -MailFrom postmaster@sedna-inc.com -MailTo exchangeadmin@sedna-inc.com -MailServer mail.sedna-inc.com

#>

[CmdletBinding()]
param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$Unit = 'GB',
    [switch]$AllExchangeServer,
    [switch]$SendMail,
    [string]$MailFrom = '',
    [string]$MailTo = '',
    [string]$MailServer = ''
)


$Unit = $Unit.ToUpper()
$now = Get-Date -Format F
$ReportTitle = ('Diskspace Report - {0}' -f ($now))
$global:Html = ''

# Set TLS protocol to TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

switch($Unit){
    'GB' {
        $ConvertTo = 1GB
    }
    'MB' {
        $ConvertTo = 1MB
    }
}

function Get-DiskspaceFromComputer {
  [CmdletBinding()]
  param(
    [string] $ServerName
  )

    if(($Unit -eq 'GB') -or ($Unit -eq 'MB')) {

        $ServerName = $ServerName.ToUpper()

        Write-Output ('Fetching Volume Data from {0}' -f ($ServerName))

        $WmiResult = Get-WmiObject Win32_Volume -ComputerName $ServerName | Select-Object Name, @{Label="Capacity ($Unit)";Expression={[decimal]::round($_.Capacity/$ConvertTo)}}, @{Label="FreeSpace ($Unit)";Expression={[decimal]::round($_.FreeSpace/$ConvertTo)}}, BootVolume, SystemVolume, FileSystem | Sort-Object Name
        $global:Html += $WmiResult | ConvertTo-Html -Fragment -PreContent ('<h2>Server {0}</h2>' -f ($ServerName))
    }

    $WmiResult
}

Function Test-SendMail {
     if( ($SendMail) -and ($MailFrom -ne '') -and ($MailTo -ne '') -and ($MailServer -ne '') ) {
        return $true
     }
     else {
        return $false
     }
}

#### MAIN
If (($SendMail) -and (!(Test-SendMail))) {
    Throw 'If -SendMail specified, -MailFrom, -MailTo and -MailServer must be specified as well!'
}

# Some CSS to get a pretty report
$head = @"

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($ReportTitle)</title>
<style type="text/css">
<!--
body {
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}
h2{ clear: both; font-size: 100%;color:#354B5E; }
h3{
    clear: both;
    font-size: 75%;
    margin-left: 20px;
    margin-top: 30px;
    color:#475F77;
}
table{
    border-collapse: collapse;
    border: none;
    font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
    color: black;
    margin-bottom: 10px;
}

table td{
    font-size: 12px;
    padding-left: 0px;
    padding-right: 20px;
    text-align: left;
}

table th {
    font-size: 12px;
    font-weight: bold;
    padding-left: 0px;
    padding-right: 20px;
    text-align: left;
}
-->
</style>
"@

if($AllExchangeServer) {
    $servers = Get-ExchangeServer | Sort-Object Name

    foreach($server in $servers) {

        $output = Get-DiskspaceFromComputer -ServerName $server.Name

        if(!($SendMail)) { $output | Format-Table -AutoSize }

    }
}
else {

    $output = Get-DiskspaceFromComputer -ServerName $ComputerName

    if(!($SendMail)) { $output | Format-Table -AutoSize }

}

if($SendMail) {

    [string]$Body = ConvertTo-Html -Body $global:Html -Title 'Status' -Head $head

    Send-MailMessage -From $MailFrom -To $Mailto -SmtpServer $MailServer -Body $Body -BodyAsHtml -Subject $ReportTitle

    Write-Output ('Email sent to {0}' -f ($MailTo))
}