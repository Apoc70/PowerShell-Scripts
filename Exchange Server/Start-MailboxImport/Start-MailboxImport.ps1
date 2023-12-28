<#
    .SYNOPSIS
    Import one or more pst files into an exisiting mailbox or a archive

    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 2.0, 2023-05-04

    .DESCRIPTION

    This script importss one or more PST files into a user mailbox or a user archive as batch.
    PST file names can used as target folder names for import. PST files are reanmed to support
    filename limitations by New-MailboxImportRequest cmdlet.

    .LINK
    https://scripts.granikos.eu

    .NOTES
    Requirements
    - Windows Server 2016+
    - Exchange Server 2016+

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial release
    1.1     log will now be stored in a subfolder (name equals Identity)
    1.2     PST file renaming added
    1.3     Module ActiveDirectory removed. We use Get-Recipient now.
    1.4		  AcceptLargeDataloss would now be added if BadItemLimit is over 51
    1.5		  Parameter IncludeFodlers added
    1.6     Parameter TargetFolder added
    1.7     Parameter Recurse added
    1.8     PST file rename after successful import added
    1.9     Updated parameter set and some PowerShell hygiene
    2.0     Updated documentation

    .PARAMETER Identity
    Type: string. Mailbox identity in which the pst files get imported

    .PARAMETER Archive
    Type: switch. Import pst files into the online archive.

    .PARAMETER FilePath
    Type:string. Folder which contains the pst files. Have to be a UNC path.

    .PARAMETER FilenameAsTargetFolder
    Type: switch. Import the pst files into targetfolders. The file name equals the target folder name.

    .PARAMETER BadItemLimit
    Type: int32. Standard is set to 0.

    .PARAMETER ContinueOnError
    Type: switch. If set the script continue with the next pst file if a import request failed.

    .PARAMETER SecondsToWait
    Type: int32. Timespan to wait between import request staus checks in seconds. Default: 320

    .PARAMETER IncludeFolders
    Type: string. If set the import would only import the given folder + subfolders. Note: If you want to import subfolders you have to use /* at the end of the folder. (Test/*).

    .PARAMETER TargetFolder
    Import the files in to definied target folder. Can't be used together with FilenameAsTargetFolder

    .PARAMETER Recurse
    If this parameter is set all PST files in subfolders will be also imported

    .PARAMETER RenameFileAfterImport
    Rename successfully imported PST files to simplify a re-run of the script. A .PST file will be renamed to .imported

    .EXAMPLE
    Import all PST file into the mailbox "testuser"
    .\Start-MailboxImport.ps1 -Idenity testuser -Filepath "\\testserver\share"

    .EXAMPLE
    Import all PST file into the mailbox "testuser". Use PST filename as target folder name. Wait 90 seconds between each status check
    .\Start-MailboxImport.ps1 -Idenity testuser -Filepath "\\testserver\share\*" -FilenameAsTargetFolder -SecondsToWait 90
#>

Param(
  [parameter(Mandatory=$true)]
    [string]$Identity,
  [parameter()]
    [switch]$Archive,
  [parameter(Mandatory=$true)]
    [string]$FilePath,
  [parameter()]
    [switch]$FilenameAsTargetFolder,
  [parameter()]
    [int]$BadItemLimit = 0,
  [parameter()]
    [switch]$ContinueOnError,
  [parameter()]
    [int]$SecondsToWait = 320,
  [parameter()]
    [string]$IncludeFolders="",
  [parameter()]
    [string]$TargetFolder="",
  [parameter()]
    [switch]$Recurse,
  [parameter()]
    [switch]$RenameFileAfterImport
)

# IMPORT GLOBAL MODULE
# Import GlobalFunctions
if($null -ne (Get-Module -Name GlobalFunctions -ListAvailable).Version) {
  Import-Module -Name GlobalFunctions
}
else {
  Write-Warning -Message 'Unable to load GlobalFunctions PowerShell module.'
  Write-Warning -Message 'Open an administrative PowerShell session and run Import-Module GlobalFunctions'
  Write-Warning -Message 'Please check http://bit.ly/GlobalFunctions for further instructions'
  exit
}
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name

# Create a log folder for each identity
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14 -LogFolder $Identity
$logger.Write('Script started')
$InfoScriptFinished = 'Script finished.'

<#
    Helpder function to remove invalid chars for New-mailboxImportRequest cmdlet
#>
Function Optimize-PstFileName {
  param (
    [Parameter(Mandatory=$true)][string]$PstFilePath
  )

  if ($Recurse) {

    # Find all files ending with .pst recursively in all folders
    $Files = Get-ChildItem -Path $PstFilePath -Include '*.pst' -Recurse

  }
  else {

    # Find all files ending with .pst in current folder
    $Files = Get-ChildItem -Path $PstFilePath -Include '*.pst'

  }

  foreach ($pst in $Files) {

    $newFileName = $pst.Name

    # List of chars, add additional chars as needed
    $chars = @(' ','(',')','&','$')
    $chars | ForEach-Object {$newFileName = $newFileName.replace($_,'')}

    $logger.Write("Renaming PST: Old: $($pst.Name) New: $($newFileName)")

    if($newFileName -ne $pst.Name) {
        # Rename PST file
        $pst | Rename-Item -NewName $newFileName
    }
  }
}

# Check if -FilenameAsTargetFolder and -TargetFolder are set both
if (($FilenameAsTargetFolder) -and ($TargetFolder)) {

  Write-Host '-FilenameAsTargetFolder and -TargetFolder can not be used together'

  Exit(1)
}

# Get all pst files from file share
if ($FilePath.StartsWith('\\')) {

  try {
    # Check file path and add wildcard, if required
    If ((!$FilePath.EndsWith('*')) -and (!$FilePath.EndsWith('\'))) {
        $FilePath = $FilePath + '\*'
    }

    Optimize-PstFileName -PstFilePath $FilePath

    # Fetch all pst files in source folder
    if ($Recurse) {
      $PstFiles = Get-ChildItem -Path $FilePath -Include '*.pst' -Recurse
    }
    else{
      $PstFiles = Get-ChildItem -Path $FilePath -Include '*.pst'
    }

    # Check if there are any files to import
    If (($PstFiles| Measure-Object).Count) {

      $InfoMessage = "Note: Script will wait $($SecondsToWait)s between each status check!"
      Write-Host $InfoMessage
      $logger.Write($InfoMessage)

      # Fetch AD user object from Active Directory
      try {
        $Name = Get-Recipient $Identity
      }
      catch {
        $InfoMessage = "Error getting recipient $($Identity). Script aborted."
        Write-Error $InfoMessage
        $logger.Write($InfoMessage, 1)
      }

      foreach ($PSTFile in $PSTFiles) {

        If ($Recurse) {
          $ImportName = $($Name.SamAccountName + '-' + $PstFiles.DirectoryName + '-' + $PstFile.Name)
        }
        else {
          $ImportName = $($Name.SamAccountName + '-' + $PstFile.Name)
        }
        $InfoMessage = "Create New-MailboxImportRequest for user: $($Name.Name) and file: $($PSTFile)"

        # Built command string
        # Checking BadItemLimit and add AcceptLargeDataLoss, if required
        if ($BadItemLimit -gt 51) {
          $cmd = "New-MailboxImportRequest -Mailbox $($($Name).SamAccountName) -Name $($ImportName) -FilePath ""$($PSTFile)"" -BadItemLimit $($BadItemLimit) -AcceptLargeDataLoss -WarningAction SilentlyContinue"
        }
        else {
          $cmd = "New-MailboxImportRequest -Mailbox $($($Name).SamAccountName) -Name $($ImportName) -FilePath ""$($PSTFile)"" -BadItemLimit $($BadItemLimit) -WarningAction SilentlyContinue"
        }

        # Checking if -Archive is set
        if ($Archive) {
          $cmd = $cmd + ' -IsArchive'
          $InfoMessage = "$($InfoMessage) into the archive."
        }
        else {
          $InfoMessage = $InfoMessage + '.'
        }

        # Check TargetFolder setup
        if ($FilenameAsTargetFolder) {
          [string]$FolderName = $($PSTFile.Name.ToString()).Replace('.pst', '')
          $cmd = $cmd + " -TargetRootFolder ""$($FolderName)"""
          $InfoMessage = $InfoMessage + " Targetfolder:""$($FolderName)""."
        }
        if ($TargetFolder) {
          $cmd = $cmd + " -TargetRootFolder ""$($TargetFolder)"""
          $InfoMessage = $InfoMessage + " Targetfolder:""$($TargetFolder)""."
        }

        # Check if IncludeFolders is set
        if ($IncludeFolders) {
          $cmd = $cmd + " -IncludeFolders ""$($IncludeFolders)"""
          $InfoMessage = $InfoMessage + " IncludeFolders:""$($IncludeFolders)""."
        }

        Write-Host $InfoMessage
        $logger.Write($InfoMessage)

        # Invoke command
        try {
          $null = Invoke-Expression -Command $cmd
        }
        catch {
          $ErrorMessage = "Error accessing creating import request for user $($Name.Name). Script aborted."

          Write-Error $ErrorMessage
          $logger.Write($ErrorMessage,1)

          Exit(1)
        }

        # Some nice sleep .zzzzzzzzzzz
        Start-Sleep -Seconds 5

        [bool]$NotFinished = $true
        $logger.Write("Waiting for import request $($ImportName) to be completed.")

        # Loop to check ongoing status of the request
        while($NotFinished) {

          try {

            $ImportRequest = Get-MailboxImportRequest -Mailbox $($($Name).SamAccountName) -Name $($ImportName) -ErrorAction SilentlyContinue

            switch ($ImportRequest.Status) {

              'Completed' {

                # Remove the ImportRequest so we can't run into the limit

                $InfoMessage = "Import request $($ImportName) completed successfully."
                Write-Host $InfoMessage
                $logger.Write("$($InfoMessage) Import Request Statistics Report:")

                # Fetch Import statistics
                $importRequestStatisticsReport = (Get-MailboxImportRequest -Mailbox $($($Name).SamAccountName) -Name $($ImportName) | Get-MailboxImportRequestStatistics -IncludeReport).Report

                # Write statistics to log, before we deleted the request (just in case need to lookup something)
                $logger.Write($importRequestStatisticsReport)

                # Rename imported PST-File if import was successful
                if ($RenameFileAfterImport) {
                  $OldFilename = Get-MailboxImportRequest -Mailbox $($($Name).SamAccountName) -Name $($ImportName) | select-object -ExpandProperty FilePath
                  Rename-Item -Path "$($OldFilename)" -NewName "$($OldFilename).imported"
                }

                # Delete mailbox import request
                Get-MailboxImportRequest -Mailbox $($($Name).SamAccountName) -Name $($ImportName) | Remove-MailboxImportRequest -Confirm:$false

                $InfoMessage = "Import request $($ImportName) deleted."
                Write-Host $InfoMessage
                $logger.Write($InfoMessage)

                $NotFinished = $false
              }

              'Failed' {

                # oops, something happend

                $InfoMessage = "Error: Administrative action is needed. ImportRequest $($ImportName) failed."

                Write-Error $InfoMessage
                $logger.Write($InfoMessage,1)

                if (-not $ContinueOnError) {
                  Write-Host $InfoScriptFinished
                  $logger.Write($InfoScriptFinished)
                  Exit(2)
                }
                else {
                  $InfoMessage = 'Info: ContinueonError is set. Continue with next PST file.'
                  Write-Host $InfoMessage
                  $logger.Write($InfoMessage)
                  $NotFinished = $false
                }
              }

              'FailedOther' {

                # oops, something special happend and we need to take care about it.

                Write-Error "Error: Administrative action is needed. ImportRequest $($ImportName) failed."
                $logger.Write("Error: Administrative action is needed. ImportRequest $($ImportName) failed.",1)

                if (-not $ContinueOnError) {
                  Write-Host $InfoScriptFinished
                  $logger.Write($InfoScriptFinished)
                  Exit(2)
                }
                else {
                  $InfoMessage = 'Info: ContinueonError is set. Continue with next pst file.'
                  Write-Host $InfoMessage
                  $logger.Write($InfoMessage)
                  $NotFinished = $false
                }
              }

              default {

                # default action: wait

                Write-Host "Waiting for import $($ImportName) to be completed. Status: $($ImportRequest.Status)"
                Start-Sleep -Seconds $SecondsToWait
              }

            }
          }
          catch {
            $InfoMessage = "Error on getting Mailboximport statistics. Trying again in $($SecondsToWait) seconds."
            Write-Host $InfoMessage
            $logger.Write($InfoMessage, 1)

            # wait before we try for the next step
            Start-Sleep -Seconds $SecondsToWait
          }
        }
      }
    }
    else {
      $InfoMessage = "No files for import found in $($FilePath)."
      Write-Host $InfoMessage
      $logger.Write($InfoMessage)
    }
  }
  catch {
    $InfoMessage = "Error accessing $($FilePath). Script aborted."
    Write-Error $InfoMessage
    $logger.Write($InfoMessage, 1)
  }
}
else {
  $InfoMessage = 'Filepath have to be an UNC path. Scipt aborted.'
  Write-Error $InfoMessage
  $logger.Write($InfoMessage, 1)
}

# Done
Write-Host $InfoScriptFinished
$logger.Write($InfoScriptFinished)