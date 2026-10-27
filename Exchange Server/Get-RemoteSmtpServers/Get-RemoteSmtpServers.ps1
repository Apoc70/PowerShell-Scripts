<#
  .SYNOPSIS
  Fetch all remote SMTP servers from Exchange receive connector logs

  Thomas Stensitzki

  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

  Version 1.3, 2023-06-17

  Ideas, comments and suggestions to support@granikos.eu

  .LINK
  http://scripts.granikos.eu

  .DESCRIPTION
  This scripts fetches remote SMTP servers by searching the Exchange receive connector logs for the EHLO string.
  Fetched servers can be exported to a single CSV file for all receive connectors across Exchange Servers or
  exported to a separate CSV file per Exchange Server.

  .NOTES
  Requirements
  - Exchange Server 2010, Exchange Server 2013+

  Revision History
  --------------------------------------------------------------------------------
  1.0 Initial community release
  1.1 Issue #2 fixed
  1.2 Minor PowerShell hygiene
  1.3 IP address, connector name, and IP uniqueness added

  .PARAMETER Servers
  List of Exchange servers, modern and legacy Exchange servers cannot be mixed

  .PARAMETER ServersToExclude
  List of host names that you want to exclude from the outout

  .PARAMETER Backend
  Search backend transport (aka hub transport) log files, instead of frontend transport, which is the default

  .PARAMETER LegacyExchange
  Search legacy Exchange servers (Exchange 2010) log file location

  .PARAMETER PathEvaluation
  Select path resolution mode. Static builds the default Exchange log path with locale-aware Program Files detection.
  Dynamic queries Exchange transport services for each server's configured receive log path.

  .PARAMETER AdminShare
  Administrative volume share used for static path evaluation, for example c$ or d$

  .PARAMETER ToCsv
  Export search results to a single CSV file for all servers

  .PARAMETER ToCsvPerServer
  Export search results to a separate CSV file per servers

  .PARAMETER UniqueIPs
  Simplify the out list by reducing the output to unique IP address

  .PARAMETER AddDays
  File selection filter, -5 will select log files changed during the last five days. Default: -10

  .EXAMPLE
  .\Get-RemoteSmtpServers.ps1 -Servers SRV01,SRV02 -LegacyExchange -AddDays -4 -ToCsv

  Search legacy Exchange servers SMTP receive log files for the last 4 days and save search results in a single CSV file

  .EXAMPLE
  .\Get-RemoteSmtpServers.ps1 -Servers SRV03,SRV04 -AddDays -4 -ToCsv -UniqueIPs

  Search Exchange servers SMTP receive log files for the last 4 days and save search results in a single CSV file, with unique IP addresses only

  .EXAMPLE
  .\Get-RemoteSmtpServers.ps1 -Servers SRV01,SRV02 -PathEvaluation Static -AdminShare c$ -Backend -ToCsv

  Search backend receive logs using a static path based on c$ and locale-aware Program Files folder detection

  .EXAMPLE
  .\Get-RemoteSmtpServers.ps1 -Servers SRV01,SRV02 -PathEvaluation Dynamic -ToCsvPerServer

  Query each server for its configured receive log path and use the derived UNC path for log parsing


#>


[CmdletBinding()]
param(
  [string[]]$Servers = @('SRV01'),
  [string[]]$ServersToExclude = @('PP2054.PPSMTP.NET', 'PP3166.PPSMTP.NET'),
  [switch]$Backend,
  [switch]$LegacyExchange,
  [ValidateSet('Static', 'Dynamic')]
  [string]$PathEvaluation = 'Static',
  [ValidatePattern('^[a-zA-Z]\$$')]
  [string]$AdminShare = 'd$',
  [switch]$ToCsv,
  [switch]$ToCsvPerServer,
  [switch]$UniqueIPs,
  [int]$AddDays = -10
)

$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path

# Set the default CSV path
$CsvFileName = Join-Path -Path $ScriptDir -ChildPath ('RemoteSMTPServers-%SERVER%-%ROLE%-{0}.csv' -f ((Get-Date).ToString('s').Replace(':', '-')))

# ToDo: Update to Get-TransportServer/Get-TransportService
# Currently pretty static
$LegacyExchangeSubPath = 'Microsoft\Exchange Server\V14\TransportRoles\Logs\ProtocolLog\SmtpReceive'
$BackendSubPath = 'Microsoft\Exchange Server\V15\TransportRoles\Logs\Hub\ProtocolLog\SmtpReceive'
$FrontendSubPath = 'Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive'

# Default Exchange Server Internal SMTP
$InternalSmtp = 'SMTP.AVAILABILITY.CONTOSO.COM'

# The SMTP receive log search pattern
$Pattern = "(.)*EHLO"

# An empty array for storing remote servers
$RemoteServers = @()
$RemoteServersList = @()

function Write-RemoteServers {
  [CmdletBinding()]
  param(
    [string]$FilePath = ''
  )

  # sort servers
  if ($UniqueIPs) {
    $RemoteServers = $RemoteServers | Sort-Object IP -Unique
  }
  else {
    $RemoteServers = $RemoteServers | Sort-Object Name
  }

  if (($RemoteServers | Measure-Object).Count -gt 0) {

    if ($ToCsv -or $ToCsvPerServer) {

      $null = $RemoteServers | Export-Csv -Path $FilePath -Encoding UTF8 -NoTypeInformation -Force -Confirm:$false

      Write-Verbose -Message ('Remote server list written to: {0}' -f $FilePath)
    }

  }
  else {
    Write-Host 'No remote servers found!'
  }

}

function Get-StaticProgramFilesFolderName {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Server,
    [string]$AdminShareName = 'd$'
  )

  $programFilesFolder = $null

  try {
    $registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Server)
    $windowsCurrentVersion = $registry.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion')
    $programFilesDir = $windowsCurrentVersion.GetValue('ProgramFilesDir')

    if (-not [string]::IsNullOrWhiteSpace($programFilesDir)) {
      $programFilesFolder = Split-Path -Path $programFilesDir -Leaf
    }
  }
  catch {
    Write-Verbose -Message ('[{0}] Unable to query ProgramFilesDir from remote registry: {1}' -f $Server, $_.Exception.Message)
  }

  if (-not $programFilesFolder) {
    $candidateFolders = @('Program Files', 'Programme', 'Programmi')

    foreach ($folderName in $candidateFolders) {
      $candidatePath = '\\{0}\{1}\{2}' -f $Server, $AdminShareName, $folderName
      if (Test-Path -Path $candidatePath) {
        $programFilesFolder = $folderName
        break
      }
    }
  }

  if (-not $programFilesFolder) {
    $programFilesFolder = 'Program Files'
  }

  return $programFilesFolder
}

function Convert-LocalPathToUncPath {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Server,
    [Parameter(Mandatory = $true)]
    [string]$LocalPath
  )

  if ($LocalPath -match '^[a-zA-Z]:\\') {
    $driveLetter = $LocalPath.Substring(0, 1)
    $pathWithoutDrive = $LocalPath.Substring(3)
    return '\\{0}\{1}$\{2}' -f $Server, $driveLetter, $pathWithoutDrive
  }

  return $null
}

function Get-ServerLogPath {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Server,
    [ValidateSet('Static', 'Dynamic')]
    [string]$EvaluationMode = 'Static',
    [string]$AdminShareName = 'd$'
  )

  if ($EvaluationMode -eq 'Dynamic') {
    $localPath = $null

    try {
      if ($LegacyExchange) {
        $localPath = (Get-TransportServer -Identity $Server -ErrorAction Stop).ReceiveProtocolLogPath.PathName
      }
      elseif ($Backend) {
        $localPath = (Get-TransportService -Identity $Server -ErrorAction Stop).ReceiveProtocolLogPath.PathName
      }
      else {
        $localPath = (Get-FrontendTransportService -Identity $Server -ErrorAction Stop).ReceiveProtocolLogPath.PathName
      }
    }
    catch {
      Write-Warning ('[{0}] Dynamic log path lookup failed: {1}' -f $Server, $_.Exception.Message)
      return $null
    }

    $remotePath = Convert-LocalPathToUncPath -Server $Server -LocalPath $localPath
    if (-not $remotePath) {
      Write-Warning ('[{0}] Unable to convert local path to UNC path: {1}' -f $Server, $localPath)
      return $null
    }

    Write-Verbose -Message ('[{0}] Dynamic local path: {1}' -f $Server, $localPath)
    Write-Verbose -Message ('[{0}] Dynamic remote path: {1}' -f $Server, $remotePath)

    return $remotePath
  }

  $programFilesFolder = Get-StaticProgramFilesFolderName -Server $Server -AdminShareName $AdminShareName

  if ($LegacyExchange) {
    $subPath = $LegacyExchangeSubPath
  }
  elseif ($Backend) {
    $subPath = $BackendSubPath
  }
  else {
    $subPath = $FrontendSubPath
  }

  $remotePath = '\\{0}\{1}\{2}\{3}' -f $Server, $AdminShareName, $programFilesFolder, $subPath

  Write-Verbose -Message ('[{0}] Static path using admin share {1}: {2}' -f $Server, $AdminShareName, $remotePath)

  return $remotePath
}

## MAIN ###########################################
# Adjust CSV file name to reflect either HUB or FRONTEND transport
if ($Backend) {
  $CsvFileName = $CsvFileName.Replace('%ROLE%', 'HUB')
}
elseif ($LegacyExchange) {
  $CsvFileName = $CsvFileName.Replace('%ROLE%', 'HUB')
}
else {
  $CsvFileName = $CsvFileName.Replace('%ROLE%', 'FE')
}

Write-Verbose -Message ('CsvFileName: {0}' -f $CsvFileName)

$ServersToExclude += $InternalSmtp

Write-Verbose -Message ($ServersToExclude | Out-String)


# Fetch each Exchange Server server
foreach ($Server in $Servers) {

  $Server = $Server.ToUpper()

  $Path = Get-ServerLogPath -Server $Server -EvaluationMode $PathEvaluation -AdminShareName $AdminShare

  if (-not $Path) {
    Write-Warning ('[{0}] Skipping server due to unresolved log path' -f $Server)
    continue
  }

  if (-not (Test-Path -Path $Path)) {
    Write-Warning ('[{0}] Resolved log path does not exist or is not accessible: {1}' -f $Server, $Path)
    continue
  }

  # fetching log files requires an account w/ administrative access to the target server
  $LogFiles = Get-ChildItem -Path $Path -File | Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays($AddDays) } # | Select-Object -First 1 # For Testing

  $LogFileCount = ($LogFiles | Measure-Object).Count
  $FileCount = 1



  foreach ($File in $LogFiles) {

    Write-Progress -Activity ('{3} | File [{0}/{1}] : {2}' -f $FileCount, $LogFileCount, $File.Name, $Server) -PercentComplete(($FileCount / $LogFileCount) * 100)

    # find results in selected log files
    $results = (Select-String -Path $File.FullName -Pattern $Pattern)

    Write-Verbose -Message ('Results {0} : {1}' -f $File.FullName, ($results | Measure-Object).Count)

    # Get remote server information from search string result
    foreach ($record in $results) {

      # Fetch Host Name
      $HostName = ($record.line -replace $Pattern, '').Replace(',', '').Trim().ToUpper()

      # Fetch IP address
      $IpAddress = ($record.line.Split(',')[5]).Split(':')[0]

      # Fetch ReceiveConnector name
      $ConnectorName = ($record.line.Split(',')[1]).Split('\')[1]

      # Add to list of host names, if not already added and not to exclude
      if (-not $RemoteServers.Contains($HostName) -and -not $ServersToExclude.Contains($HostName) -and -not ($IpAddress -eq '127.0.0.1')) {

        $RemoteServersList += ('{0} [{1}]' -f $HostName, $IpAddress)

        # Build custom object and store data
        $RemoteServer = New-Object PSCustomObject

        $RemoteServer | Add-Member -Type NoteProperty -Name 'Name' -Value $Hostname
        $RemoteServer | Add-Member -Type NoteProperty -Name 'IP' -Value $IpAddress
        $RemoteServer | Add-Member -Type NoteProperty -Name 'ConnectorName' -Value $ConnectorName

        $RemoteServers += $RemoteServer

      }
    }

    $FileCount++

  }

  if ($ToCsvPerServer) {

    # Export a signkle CSV file per queried server
    $CsvFile = $CsvFileName.Replace('%SERVER%', $Server)

    Write-Verbose -Message $CsvFile

    Write-RemoteServers -FilePath $CsvFile

    $RemoteServers = @()
  }
}

if ($ToCsv) {

  # Export a single CSV file
  $CsvFile = $CsvFileName.Replace('%SERVER%', 'ALL')

  Write-Verbose -Message $CsvFile

  Write-RemoteServers -FilePath $CsvFile
}