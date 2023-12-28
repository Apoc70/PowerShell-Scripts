<#
    .SYNOPSIS
    Fetch all remote SMTP servers from Exchange receive connector logs, establishing a TLS connection
   
    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.0, 2021-06-04

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    http://scripts.granikos.eu

    .DESCRIPTION
    This scripts fetches remote SMTP servers by searching the Exchange receive connector logs for successful TLS connections.
    Fetched servers can be exported to a single CSV file for all receive connectors across Exchange Servers or
    exported to a separate CSV file per Exchange Server.
    You can use this script to identify remote servers connecting using TLS 1.0 or TLS 1.1.

	
    .NOTES 
    Requirements 
    - Exchange Server 2013+

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
	
    .PARAMETER Servers
    List of Exchange servers, modern and legacy Exchange servers cannot be mixed

    .PARAMETER Backend
    Search backend transport (aka hub transport) log files, instead of frontend transport, which is the default

    .PARAMETER ToCsv
    Export search results to a single CSV file for all servers

    .PRAMATER ToCsvPerServer
    Export search results to a separate CSV file per servers

    .PARAMETER AddDays
    File selection filter, -5 will select log files changed during the last five days. Default: -10

  
    .EXAMPLE
    .\Get-RemoteSmtpServers.ps1 -Servers SRV01,SRV02 -LegacyExchange -AddDays -4 -ToCsv

    Search legacy Exchange servers SMTP receive log files for the last 4 days and save search results in a single CSV file
   
#>


[CmdletBinding()]
param(
  $Servers = @('MYEXCHANGE'),
  [switch]$Backend,
  [switch]$ToCsv,
  [switch]$ToCsvPerServer,
  [int]$AddDays = -10
)

$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path

$CsvFileName = ('RemoteSMTPServersTls-%SERVER%-%ROLE%-%TLS%-{0}.csv' -f ((Get-Date).ToString('s').Replace(':','-')))

# ToDo: Update to Get-TransportServer/Get-TransportService 
# Currently pretty static
$BackendPath = '\\%SERVER%\d$\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Hub\ProtocolLog\SmtpReceive'
$FrontendPath = '\\%SERVER%\d$\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive'

# The TLS version to search
$TlsProtocols = @('TLS1_2','TLS1_1','TLS1_0')

# The SMTP receive log search pattern
$Pattern = '(.)*SP_PROT_%TLS%(.)*succeeded'

# An empty array for storing remote servers
$RemoteServers = @()

# Create an empty array
$RemoteServersOutput = @()

function Write-RemoteServers {
  [CmdletBinding()]
  param(
    [string]$FilePath = ''
  )

  # sort servers
  $RemoteServers = $RemoteServers | Select-Object -Unique | Sort-Object

  if(($RemoteServers| Measure-Object).Count -gt 0) {
    
    # Create an empty array
    $RemoteServersOutput = @()

    foreach($Server in $RemoteServers) { 
    
      if($Server.Trim() -ne '') { 
        $obj = New-Object -TypeName PSObject
        $obj | Add-Member -MemberType NoteProperty -Name 'Remote Server' -Value $Server
        $RemoteServersOutput += $obj
      }
    }

    if($ToCsv -or $ToCsvPerServer) {
      # save remote servers list as csv
      $null = $RemoteServersOutput | Export-Csv -Path $FilePath -Encoding UTF8 -NoTypeInformation -Force -Confirm:$false

      Write-Verbose -Message ('Remote server list written to: {0}' -f $FilePath)
    }

    $RemoteServersOutput
  
  }
  else {
    Write-Host 'No remote servers found!' 
  }

}

## MAIN ###########################################
$LogPath = $FrontendPath

# Adjust CSV file name to reflect either HUB or FRONTEND transport
if($Backend) {
  $LogPath = $BackendPath
  $CsvFileName = $CsvFileName.Replace('%ROLE%','HUB')
}
else {
  $CsvFileName = $CsvFileName.Replace('%ROLE%','FE')
}

Write-Verbose -Message ('CsvFileName: {0}' -f ($CsvFileName))

# Fetch each Exchange Server server 
foreach($Server in $Servers) {
 
  $Server = $Server.ToUpper()

  $Path = $LogPath.Replace('%SERVER%', $Server)

  Write-Verbose -Message ('Working on Server {0} | {1}' -f $Server, $Path)

  # fetching log files requires an account w/ administrative access to the target server
  $LogFiles = Get-ChildItem -Path $Path -File | Where-Object {$_.LastWriteTime -gt (Get-Date).AddDays($AddDays)} | Select-Object -First 2

  $LogFileCount = ($LogFiles | Measure-Object).Count
  
  foreach($Tls in $TlsProtocols) {
    Write-Host ('Working on: {0}' -f $Tls)
    $FileCount = 1

    foreach($File in $LogFiles) {

      Write-Progress -Activity ('{3} | {4} | File [{0}/{1}] : {2}' -f $FileCount, $LogFileCount, $File.Name, $Server, $Tls) -PercentComplete(($FileCount/$LogFileCount)*100)

 
      # find results in selected log files
      $SearchPattern = $Pattern.Replace('%TLS%', $Tls)

      $results = (Select-String -Path $File.FullName -Pattern $SearchPattern)

      Write-Verbose -Message ('Results {0} : {1}' -f $File.FullName, ($results | Measure-Object).Count)

      # Get remote server information from search string result
      foreach($record in $results) {
        
        # Fetch remote hostname
        # $HostName = ($record.line -replace $Pattern,'').Replace(',','').Trim().ToUpper()
        $HostIp = (($record.Line).Split(',')[5]).Split(':')[0]

        # Try to resolve remote IP address as the line does not contain a server name
        $HostName = Resolve-DnsName $HostIp -ErrorAction Ignore |Select-Object -ExpandProperty NameHost

        if(-not $RemoteServers.Contains($HostName)) { 
          $RemoteServers += $HostName
        }
      }

      $FileCount++

    }

    if($ToCsvPerServer) {
   
      $CsvFile = $CsvFileName.Replace('%SERVER%',$Server).Replace('%TLS%',$Tls)

      Write-Verbose -Message $CsvFile

      Write-RemoteServers -FilePath $CsvFile
    
      $RemoteServers = @()
    }
  }
}

if($ToCsv) { 
  $CsvFile = $CsvFileName.Replace('%SERVER%','ALL')

  Write-Verbose -Message $CsvFile

  Write-RemoteServers -FilePath $CsvFile
}