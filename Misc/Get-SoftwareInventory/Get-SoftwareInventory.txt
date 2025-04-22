<#
  .SYNOPSIS

  Script to generate an html reports of installed software, installed updates and installed components on a remote computer

  .LINK

  https://script.granikos.eu

  .DESCRIPTION

  This script utilizes the slightly modified function Get-RemoteProgram by Jaap Brasser to fetch information about installed software.

  The script generates outputs for
  * Installed programs
  * Installed updates
  * Installed components

  These three outputs are composed to a single html emails. All three outputs are
  exported to CSV files which are added to the generated email as attachments. The
  html report is added as attachment as well.

  Written by: Thomas Stensitzki

  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

  Version 1.1, 2025-03-13

  Please send ideas, comments and suggestions to support@granikos.eu

  .NOTES

  Requirements
  - Windows Server 2019+
  - Remote Registry access to remote computers

  Revision History
  --------------------------------------------------------------------------------
    1.0      Initial community release
    1.1      Added support for filtering installed software

  .CREDITS
  Jaap Brasser
  Georges Zwingelstein

  .PARAMETER RemoteComputer
  Name of the computer to fetch data from (default: local computer)

  .PARAMETER SendMail
  Switch to send an Html report

  .PARAMETER MailFrom
  Email address of report sender

  .PARAMETER MailTo
  Email address of report recipient

  .PARAMETER MailServer
  SMTP Server for email report

  .PARAMETER SoftwareFilter
  Filter for installed software. Only software matching the filter will be included in the report.

  .PARAMETER ProgramsOnly
  Switch to only fetch installed programs and exclude updates and components

  .EXAMPLE
  Get software information about SERVER01, SERVER02 and send email report

  .\Get-SoftwareInventory.ps1 -SendMail -MailFrom postmaster@mcsmemail.de -MailTo it-support@mcsmemail.de -MailServer mymailserver.mcsmemail.de -RemoteComputer SERVER01,SERVER02

  .EXAMPLE
  Get software information about SERVER01, SERVER02 and write report to disk only

  .\Get-SoftwareInventory.ps1 -RemoteComputer SERVER01,SERVER02
#>

[CmdletBinding()]
param(
    [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
    [string[]]
    $RemoteComputer = $env:COMPUTERNAME,
    [string]$SoftwareFilter = '',
    [switch]$ProgramsOnly,
    [switch]$SendMail,
    [string]$MailFrom = '',
    [string]$MailTo = '',
    [string]$MailServer = ''
)

$ScriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Path

$now = Get-Date -Format 'yyyy-MM-dd HH:mm'
$ReportTitle = 'Software Inventory Report'
$count = 1

<#
    .NOTES
    Name: Get-RemoteProgram
    Author: Jaap Brasser
    Version: 1.2.1
    DateCreated: 2013-08-23
    DateUpdated: 2015-02-28
    Blog: http://www.jaapbrasser.com

    .LINK
    http://www.jaapbrasser.com

    .LINK
    https://gallery.technet.microsoft.com/scriptcenter/Get-RemoteProgram-Get-list-de9fd2b4
#>
Function Get-RemoteProgram {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [string[]]
        $ComputerName = $env:COMPUTERNAME,
        [Parameter(Position = 0)]
        [string[]]
        $Property,
        [switch]
        $ExcludeSimilar,
        [int]
        $SimilarWord
    )

    begin {
        $RegistryLocation = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\',
        'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
        $HashProperty = @{}
        $SelectProperty = @('ProgramName', 'ComputerName')
        if ($Property) {
            $SelectProperty += $Property
        }
    }

    process {
        foreach ($Computer in $ComputerName) {
            $RegBase = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Computer)
            $RegistryLocation | ForEach-Object {
                $CurrentReg = $_
                if ($RegBase) {
                    $CurrentRegKey = $RegBase.OpenSubKey($CurrentReg)
                    if ($CurrentRegKey) {
                        $CurrentRegKey.GetSubKeyNames() | ForEach-Object {
                            if ($Property) {
                                foreach ($CurrentProperty in $Property) {
                                    $HashProperty.$CurrentProperty = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue($CurrentProperty)
                                }
                            }
                            $HashProperty.ComputerName = $Computer
                            $HashProperty.ProgramName = ($DisplayName = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue('DisplayName'))
                            if ($DisplayName) {
                                New-Object -TypeName PSCustomObject -Property $HashProperty |
                                Select-Object -Property $SelectProperty
                            }
                        }
                    }
                }
            } | ForEach-Object -Begin {
                if ($SimilarWord) {
                    $Regex = [regex]"(^(.+?\s){$SimilarWord}).*$|(.*)"
                }
                else {
                    $Regex = [regex]"(^(.+?\s){3}).*$|(.*)"
                }
                [System.Collections.ArrayList]$Array = @()
            } -Process {
                if ($ExcludeSimilar) {
                    $null = $Array.Add($_)
                }
                else {
                    $_
                }
            } -End {
                if ($ExcludeSimilar) {
                    $Array | Select-Object -Property *, @{
                        name       = 'GroupedName'
                        expression = {
                            ($_.ProgramName -split $Regex)[1]
                        }
                    } |
                    Group-Object -Property 'GroupedName' | ForEach-Object {
                        $_.Group[0] | Select-Object -Property * -ExcludeProperty GroupedName
                    }
                }
            }
        }
    }
}

### MAIN ######################################################
# Query string courtesy of Georges Zwingelstein

# Some CSS to create a pretty report
$head = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($ReportTitle)</title>
<style type="text/css">
<!-
body {
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}
h1{ clear: both;
    font: 1.2em sans-serif;
    color:#000000;
}
h2{
    clear: both;
    font: 1.1em sans-serif;
    color:#354B5E;
}
h3{
    clear: both;
    font-size: 75%;
    margin-left: 20px;
    margin-top: 30px;
    color:#475F77;
}
table{
    border-collapse: collapse;
    border: solid;
    font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
    color: black;
    margin-bottom: 10px;
}

table tr {
    border: solid;
}

table td{
    font-size: 12px;
    padding-left: 5px;
    padding-right: 5px;
    text-align: left;
}

table th {
    font-size: 12px;
    font-weight: bold;
    padding-left: 5px;
    padding-right: 5px;
    text-align: left;

}
->
</style>
"@

foreach ($Computer in $RemoteComputer) {

    $Computer = $Computer.ToUpper()

    Write-Progress -Activity 'Fetching software inventory' -Status ('Working on server {0} [{1}/{2}]' -f $computer, $count, $RemoteComputer.Count) -PercentComplete (($count / $RemoteComputer.Count) * 100)

    $HtmlReport = @()
    $Attachments = @()
    $ReportTitle = ('Software Inventory Report [{0}] - {1}' -f $computer.ToUpper(), $now)
    # $csvFile = ''

    # Fetch installed programs
    $InstalledPrograms = Get-RemoteProgram -ComputerName $computer -Property Publisher, InstallDate, DisplayVersion, IsMinorUpgrade, ReleaseType, ParentDisplayName, SystemComponent | Where-Object { [string]$_.SystemComponent -ne 1 -and ![string]$_.IsMinorUpgrade -and ![string]$_.ReleaseType -and ![string]$_.ParentDisplayName } | Sort-Object -Property ProgramName

    if ($SoftwareFilter) {
        $InstalledPrograms = $InstalledPrograms | Where-Object { $_.ProgramName -like "*$SoftwareFilter*" }
    }

    if ($null -ne $InstalledPrograms) {
        $csvInstalledPrograms = ('{0}\InstalledPrograms-{1}.csv' -f $ScriptDir, $computer.ToUpper())
        $InstalledPrograms | Export-Csv -Path $csvInstalledPrograms -Force -Encoding UTF8 -NoTypeInformation -Delimiter ';'
        $HtmlReport += $InstalledPrograms | ConvertTo-Html -Fragment -PreContent ('<h1>{0}</h1><h2>Installed Programs</h2>' -f $computer.ToUpper())
        $Attachments += $csvInstalledPrograms
    }
    else {
        $HtmlReport += ('<h1>{0}</h1><h2>Installed Programs</h2><p>Error fetching installed programs</p>' -f $computer.ToUpper())
    }

    if (-not $ProgramsOnly) {

        # Fetch installed updates
        $InstalledUpdates = Get-RemoteProgram -ComputerName $computer -Property Publisher, InstallDate, DisplayVersion, IsMinorUpgrade, ReleaseType, ParentDisplayName, SystemComponent | Where-Object { [string]$_.SystemComponent -ne 1 -and ([string]$_.IsMinorUpgrade -or [string]$_.ReleaseType -or [string]$_.ParentDisplayName) } | Sort-Object -Property ParentDisplayName, ProgramName

        if ($null -ne $InstalledUpdates) {
            $csvInstalledUpdates = ('{0}\InstalledUpdates-{1}.csv' -f $ScriptDir, $computer.ToUpper())
            $InstalledUpdates | Export-Csv -Path $csvInstalledUpdates -Force -Encoding UTF8 -NoTypeInformation -Delimiter ';'
            $HtmlReport += $InstalledUpdates | ConvertTo-Html -Fragment -PreContent '<h2>Installed Updates</h2>'
            $Attachments += $csvInstalledUpdates
        }
        else {
            $HtmlReport += ('<h1>{0}</h1><h2>Installed Updates</h2><p>Error fetching installed updates</p>' -f $computer.ToUpper())
        }

        # Fetch installed components
        $InstalledComponents = Get-RemoteProgram -ComputerName $computer -Property Publisher, InstallDate, DisplayVersion, IsMinorUpgrade, ReleaseType, ParentDisplayName, SystemComponent | Where-Object { [string]$_.SystemComponent -eq 1 } | Sort-Object -Property ProgramName

        if ($null -ne $InstalledComponents) {
            $csvInstalledComponents = ('{0}\InstalledComponents-{1}.csv' -f $ScriptDir, $computer.ToUpper())
            $InstalledComponents | Export-Csv -Path $csvInstalledComponents -Force -Encoding UTF8 -NoTypeInformation -Delimiter ';'
            $HtmlReport += $InstalledComponents | ConvertTo-Html -Fragment -PreContent '<h2>Installed Components</h2>'
            $Attachments += $csvInstalledComponents
        }
        else {
            $HtmlReport += ('<h1>{0}</h1><h2>Installed Components</h2><p>Error fetching installed components</p>' -f $computer.ToUpper())
        }
    }

    # Generate full html report
    [string]$HtmlReport = ConvertTo-Html -Body $HtmlReport -Head $head -Title $ReportTitle
    $reportFile = ('{0}\SoftwareInventory-{1}.html' -f $ScriptDir, $computer.ToUpper())
    $HtmlReport | Out-File -Encoding utf8 -FilePath $reportFile -Force
    $Attachments += $reportFile

    # Send email report
    if ($SendMail) {
        Send-MailMessage -From $MailFrom -To $MailTo -SmtpServer $MailServer -Body $HtmlReport -Subject $ReportTitle -Attachments $Attachments -BodyAsHtml
    }
    $count++
}

Write-Progress -Activity 'Fetching software inventory' -Status 'Completed' -Completed

Return 0