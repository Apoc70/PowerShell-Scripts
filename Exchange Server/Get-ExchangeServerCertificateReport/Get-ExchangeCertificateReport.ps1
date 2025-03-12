<#
.SYNOPSIS

Exchange Server SSL Certificate Report Script

.DESCRIPTION

Generates a report of the SSL certificates installed on Exchange Server servers

.OUTPUTS

Outputs to a HTML file.

.EXAMPLE

.\CertificateReport.ps1
Reports SSL certificates for Exchange Server servers and outputs to a HTML file.

.LINK

http://exchangeserverpro.com/powershell-script-ssl-certificate-report (offline)

.NOTES

Written By: Paul Cunningham
Website:	http://exchangeserverpro.com
Twitter:	http://twitter.com/exchservpro

Updated By: Thomas Stensitzki

Change Log
V1.00, 13/03/2014 - Initial Version
V1.01, 13/03/2014 - Minor bug fix
V2.00, 2025-03-12 - Some PowerShell optimizations, changed email sending to Send-MailMessage

#>
[CmdletBinding()]
param(
    [parameter(Mandatory = $false, HelpMessage = 'Send report as Html email')]
    [switch] $SendMail,
    [parameter(Mandatory = $false, HelpMessage = 'Sender address for result summary')]
    [string]$MailFrom = '',
    [parameter(Mandatory = $false, HelpMessage = 'Recipient address for result summary')]
    [string]$MailTo = '',
    [parameter(Mandatory = $false, HelpMessage = 'SMTP Server address for sending result summary')]
    [string]$MailServer = '',
    [string]$cssFilenname = 'styles.css'
)

$scriptVersion = '2.00'

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

$reportFile = ('{0}\CertificationReport-{1}.html' -f $scriptPath, (Get-Date -Format yyyyMMdd-HHmm))
$cssFile = Join-Path -Path $scriptPath -ChildPath $cssFilenname

$now = Get-Date -Format F

$reportTitle = ('Exchange Server Certificate Report - {0}' -f $now)

Function Check-SendMail {
    if ( ($SendMail) -and ($MailFrom -ne '') -and ($MailTo -ne '') -and ($MailServer -ne '') ) {
        return $true
    }
    else {
        return $false
    }
}

#### MAIN
If (($SendMail) -and (!(Check-SendMail))) {
    Throw "If -SendMail specified, -MailFrom, -MailTo and -MailServer must be specified as well!"
}

$htmlReport = @()
$exchangeServers = @(Get-ExchangeServer in* | Sort-Object Name)

foreach ($server in $exchangeServers) {
    # $htmlsegment = @()

    $serverDetails = ('Server: {0} ({1})' -f $server.Name, $server.ServerRole)
    Write-Output $serverDetails

    $certTable = @()

    try {
        $certificates = $null
        $certificates = @(Get-ExchangeCertificate -Server $server -ErrorAction SilentlyContinue ) #| Sort-Object NotAfter)

        if (($certificates | Measure-Object).Count -ne 0) {

            Write-Verbose ('{0} certificates found...' -f $certificates.Count)

            # Check each certificate
            foreach ($cert in $certificates) {

                # variables for Exchange assigned Exchange services
                $iis = $null
                $smtp = $null
                $pop = $null
                $imap = $null
                $um = $null

                $subject = ((($cert.Subject -split ",")[0]) -split "=")[1]

                Write-Verbose "Subject: $($subject)"

                if ($($cert.IsSelfSigned)) {
                    $selfsigned = "Yes"
                }
                else {
                    $selfsigned = "No"
                }

                $issuer = ((($cert.Issuer -split ",")[0]) -split "=")[1]

                $domains = ''
                $certDomains = @($cert | Select-Object -ExpandProperty:CertificateDomains)

                if ($($certDomains.Count) -gt 1) {
                    $domains = $null
                    $domains = $certDomains -join ", "
                }
                elseif (($certDomains | Measure-Object).Count -ne 0) {
                    $domains = $certDomains[0]
                }

                $services = $cert.ServicesStringForm.ToCharArray()

                if ($services -icontains "W") { $iis = "Yes" }
                if ($services -icontains "S") { $smtp = "Yes" }
                if ($services -icontains "P") { $pop = "Yes" }
                if ($services -icontains "I") { $imap = "Yes" }
                if ($services -icontains "U") { $um = "Yes" }

                # Create a new object for the certificate
                $certObj = New-Object PSObject
                $certObj | Add-Member NoteProperty -Name "Subject" -Value $subject
                $certObj | Add-Member NoteProperty -Name "Status" -Value $cert.Status
                $certObj | Add-Member NoteProperty -Name "Expires" -Value $cert.NotAfter.ToShortDateString()
                $certObj | Add-Member NoteProperty -Name "Self Signed" -Value $selfsigned
                $certObj | Add-Member NoteProperty -Name "Issuer" -Value $issuer
                $certObj | Add-Member NoteProperty -Name "SMTP" -Value $smtp
                $certObj | Add-Member NoteProperty -Name "IIS" -Value $iis
                $certObj | Add-Member NoteProperty -Name "POP" -Value $pop
                $certObj | Add-Member NoteProperty -Name "IMAP" -Value $imap
                $certObj | Add-Member NoteProperty -Name "UM" -Value $um
                $certObj | Add-Member NoteProperty -Name "Thumbprint" -Value $cert.Thumbprint
                $certObj | Add-Member NoteProperty -Name "Domains" -Value $domains

                # Add the certificate object to the certificate table
                $certTable += $certObj
            }
        }
    }
    catch {
        Write-Output ('Error accessing server {0}' -f $server.Name)
        Write-Output $_.Exception.Message
     }

    if (($certTable | Measure-Object).Count -ne 0) {
        $html = $certTable.GetEnumerator() | Sort-Object -Property Subject, Expires | ConvertTo-Html -Fragment -PreContent "<h2>$serverDetails</h2>"

        $htmlReport += $html #| ConvertTo-Html -Fragment -PreContent "<h2>$serverDetails</h2>"
    }
    else {
        $htmlReport += "<h2>$serverDetails</h2><p class=''error''>Some error occured while trying to access this server.</p>"
    }
}
$htmlReport += ('<p>Script Version: {0}</p>' -f $scriptVersion)

$htmlReport += "<p><a class= ''link'' href='https://learn.microsoft.com/exchange/architecture/client-access/certificates?view=exchserver-2019'>Certificates in Exchange Online Documentation</a></p>"


$head = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
<head><title>$($reportTitle)</title>
<style type="text/css">$(Get-Content $cssFile)</style></head>
<body><h1 align=''center''>$($reportTitle)</h1>
"@

[string]$htmlReport = ConvertTo-Html -Body $htmlReport -Head $head -Title $reportTitle

$htmlReport | Out-File -Encoding utf8 -FilePath $reportfile -Force

# Send certificate report as email
if ($SendMail) {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;
    Send-MailMessage -From $MailFrom -To $MailTo -SmtpServer $MailServer -Body $htmlReport -BodyAsHtml -Subject $reportTitle -Encoding ([System.Text.Encoding]::UTF8)
}

Return 0