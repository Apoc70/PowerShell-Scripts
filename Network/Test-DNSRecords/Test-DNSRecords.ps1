param(
    [string]$DomainCsvFile = '.\Domains.txt',
    [switch]$SendMail,
    [switch]$ExportToCsv,
    [switch]$TestExoDkim,
    [switch]$ResolveSpfIncludes,
    $DNSServer = '8.8.8.8'
)

# https://powershellisfun.com/2022/06/19/retrieve-email-dns-records-using-powershell/

$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$Delimiter = ';'

# Your SMTP server relay information
$SmtpServer = "smtp.yourserver.com"
$SmtpPort = 25
$To = "recipient@domain.com"
$From = "yourEmail@domain.com"
$MessageSubject = "DNS Mail Records Report"


function Export-TestResultToCsv {
    param(

        $ResultSetToExport
    )

    $FilePath = Join-Path -Path $ScriptDir -ChildPath ('DNSTestResult-{0}.csv' -f (Get-Date).ToString('yyyyMMdd') )

    Write-Output ('Exporting results to {0}' -f $FilePath)

    $ResultSet | Export-Csv -Path $FilePath -Delimiter $Delimiter -Encoding utf8 -Force -Confirm:$false

    return $FilePath
}

if (Test-Path -Path $DomainCsvFile) {

    $Domains = Import-Csv -Path $DomainCsvFile

}

if ($null -ne $Domains) {

    Write-Output ('Using DNS {0} to resolve {1} domain(s)' -f $DNSServer, ($Domains | Measure-Object).Count)

    $ResultSet = @()

    Write-Output ('')

    foreach ($DomainEntry in $Domains) {

        #Check if domain name is valid, output warning it not and continue to the next domain (if any)
        try {
            $domain = $DomainEntry.Domain

            Write-Output ('Resolving: {0}' -f $domain)
            Resolve-DnsName -Name $domain -Server $DNSserver -ErrorAction Stop | Out-Null

            # Test AutoDiscover Ad recrod
            $AutoDiscoverARecord = (Resolve-DnsName -Name ('autodiscover.{0}' -f$domain) -Type A -Server $DNSserver -ErrorAction SilentlyContinue).IPAddress
            $AutoDiscoverCNAMERecord = (Resolve-DnsName -Name ('autodiscover.{0}' -f$domain) -Type CNAME -Server $DNSserver -ErrorAction SilentlyContinue).NameHost

            # DKIM
            $DKIM1Record = Resolve-DnsName -Name ('selector1._domainkey.{0}' -f $domain) -Type CNAME -Server $DNSserver -ErrorAction SilentlyContinue
            $DKIM2Record = Resolve-DnsName -Name ('selector1._domainkey.{0}' -f $domain) -Type CNAME -Server $DNSserver -ErrorAction SilentlyContinue

            # DMARCRecord
            $DMARCRecord = (Resolve-DnsName -Name ('_dmarc.{0}' -f $domain) -Type TXT -Server $DNSserver -ErrorAction SilentlyContinue | Where-Object Strings -Match 'DMARC').Strings

            # MTA-STS
            $MTASTSRecord = (Resolve-DnsName -Name ('_mta-sts.{0}' -f $domain) -Type TXT -Server $DNSserver -ErrorAction SilentlyContinue | Where-Object Strings -Match 'v=sts').Strings

            #MX
            $MXRecord = (Resolve-DnsName -Name $domain -Type MX -Server $DNSserver -ErrorAction SilentlyContinue).NameExchange

            # SPF
            $SPFRecord = (Resolve-DnsName -Name $domain -Type TXT -Server $DNSserver -ErrorAction SilentlyContinue | Where-Object Strings -Match 'v=spf').Strings
            $includes = (Resolve-DnsName -Name $domain -Type TXT -Server $DNSserver -ErrorAction SilentlyContinue | Where-Object Strings -Match 'v=spf').Strings -split ' ' | Select-String 'Include:'

            #Set variables to Not enabled or found if they can't be retrieved
            $ErrorText = 'Not configured'


            if ($null -eq $DKIM1Record -and $null -eq $DKIM2Record) {
                $dkim = ('{0} or using custom DKIM hostname' -f $ErrorText)
            }
            else {
                $dkim = "$($DKIM1Record.Name) , $($DKIM2Record.Name)"
            }

            if ($null -eq $MTASTSRecord) {
                $MTASTSRecord = $ErrorText
            }
            else {
                # Fetch MTA-STS text file
                $MTASTSContent = Invoke-WebRequest -Uri ('https://mta-sts.{0}/.well-known/mta-sts.txt' -f $domain) -ErrorAction SilentlyContinue
                if($null -ne $MTASTSContent) {
                    $MTASTSTextFile = $MTASTSContent.Content.Replace([System.Environment]::NewLine,', ')
                }
                else {
                    $MTASTSTextFile = $ErrorText
                }

            }

            if ($null -eq $DMARCRecord) {
                $DMARCRecord = $ErrorText
            }

            if ($null -eq $MXRecord) {
                $MXRecord = $ErrorText
            }

            if ($null -eq $SPFRecord) {
                $SPFRecord = $ErrorText
            }

            if ($null -eq $AutoDiscoverCNAMERecord) {
                $AutoDiscoverCNAMERecord = $ErrorText
            }

            if (($AutoDiscoverARecord).count -gt 1 -or $null -ne $AutoDiscoverCNAMERecord) {
                $AutoDiscoverARecord = $ErrorText
            }

            if ($null -eq $includes) {
                $includes = $ErrorText
            }
            else {
                $foundincludes = foreach ($include in $includes) {
                    if ((Resolve-DnsName -Server $DNSserver -Name $include.ToString().Split(':')[1] -Type txt -ErrorAction SilentlyContinue).Strings) {
                        [PSCustomObject]@{
                            SPFIncludes = "$($include.ToString().Split(':')[1]) : " + $(Resolve-DnsName -Server $DNSserver -Name $include.ToString().Split(':')[1] -Type txt).Strings
                        }
                    }
                    else {
                        [PSCustomObject]@{
                            SPFIncludes = $ErrorText
                        }
                    }
                }
            }

            $ResultSet += [PSCustomObject]@{
                'Domain Name'             = $domain
                'Autodiscover IP-Address' = $AutoDiscoverARecord
                'Autodiscover CNAME '     = $AutoDiscoverCNAMERecord
                'DKIM Record'             = $dkim
                'DMARC Record'            = "$($DMARCRecord)"
                'MTA-STS Record'          = "$($MTASTSRecord)"
                'MTA-STS File Content'    = $MTASTSTextFile
                'MX Record(s)'            = $MXRecord -join ', '
                'SPF Record'              = "$($SPFRecord)"
                #'SPF Include values'      = "$($foundincludes.SPFIncludes)" -replace "all", "all`n`b"
                'SPF Include values'      = "$($foundincludes.SPFIncludes)" #-replace "all", "all"
            }

        }
        catch{}
    }

    if ($ExportToCsv) {
        # Simply export to a CSV file in the script directory
        Export-TestResultToCsv -ResultSetToExport $ResultSet
    }

    if ($SendMail) {

        # Export results to CSV and fetch file path for attachment
        $CsvFilePath = Export-TestResultToCsv -ResultSetToExport $ResultSet

        # Send email
        $emailBody = $ResultSet | ConvertTo-Html | Out-String

        $smtpServer = "smtp.yourserver.com"
        $smtpPort = 587
        $smtpUser = "yourEmail@domain.com"
        $smtpPass = "yourPassword"
        $to = "recipient@domain.com"
        $from = "yourEmail@domain.com"
        $subject = "DNS Mail Records Report"

        $credentials = New-Object System.Management.Automation.PSCredential ($smtpUser, (ConvertTo-SecureString $smtpPass -AsPlainText -Force))

        Send-MailMessage -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $credentials -To $to -From $from -Subject $subject -Body $emailBody -BodyAsHtml -Attachments $CsvFilePath

    }
    else {
        Write-Output $ResultSet
    }
}