<#
    .SYNOPSIS
    Export messages from a transport queue to file system for manual replay

    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.2, 2022-09-05

    Leave ideas, comments, and suggestions at GitHub

    .LINK
    https://github.com/Apoc70/Export-MessageQueue

    .DESCRIPTION

    This script suspends a transport queue, exports the messages to the local file system. After successful export the messages are optionally deleted from the queue.

    .NOTES
    Requirements
    - Exchange Server 2016+
    - Windows Server 2016+
    - Utilizes global functions library

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial community release
    1.1     Some PowerShell hygiene
    1.2     PowerSHell Typo fixed and updated for Exchange Server 2019

    .PARAMETER Queue
    Full name of the transport queue, e.g. SERVER\354
    Use Get-Queue -Server SERVERNAME to identify message queue

    .PARAMETER Path
    Path to folder for exprted messages

    .PARAMETER DeleteAfterExport
    Switch to delete per Exchange Server subfolders and creating new folders

    .EXAMPLE
    Export messages from queue MCMEP01\45534 to D:\ExportedMessages and do not delete messages after export

    .\Export-MessageQueue -Queue MCMEP01\45534 -Path D:\ExportedMessages

    .EXAMPLE
    Export messages from queue MCMEP01\45534 to D:\ExportedMessages and delete messages after export

    .\Export-MessageQueue -Queue MCMEP01\45534 -Path D:\ExportedMessages -DeleteAfterExport

#>
param(
  [parameter(Mandatory=$true,HelpMessage='Transport queue holding messages to be exported (e.g. SERVER\354)')]
    [string] $Queue,
  [parameter(Mandatory=$true,HelpMessage='File path to local folder for exprted messages (e.g. E:\Export)')]
    [string] $Path,
    [switch] $DeleteAfterExport
)

# Set-StrictMode -Version Latest

# Implementation of global module
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

$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write('Script started')
$logger.Write(('Working on message queue {0}, Export folder: {1}, DeleteAfterExport: {2}' -f $Queue, $Path, $DeleteAfterExport))

### FUNCTIONS -----------------------------

function Request-Choice {
    [CmdletBinding()]
    param([string]$Caption)
    $choices =  [System.Management.Automation.Host.ChoiceDescription[]]@("&Yes","&No")
    [int]$defaultChoice = 1

    $choiceReturn = $Host.UI.PromptForChoice($Caption, "", $choices, $defaultChoice)

    return $choiceReturn
}

function Check-Folders {
    # Check, if export folder exists
    if(!(Test-Path -Path $Path)) {
        # Folder does not exist, lets create a new root folder
        New-Item -Path $Path -ItemType Directory | Out-Null

        $logger.Write(('Folder {0} created' -f $Path))
    }
}


function Check-Queue {
    # Check message queue
    $messageCount = -1
    try {
        $messageQueue = Get-Queue $Queue
        $messageCount = $messageQueue.MessageCount

        $logger.Write(('{0} message(s) found in queue {1}' -f $messageCount, $Queue))
    }
    catch {
        $logger.Write(('Queue {0} cannot be accessed' -f $Queue))
    }
    $messageCount
}

function Export-Messages {
    # Export suspended messages
    try {
        # Suspend messages in queue
        $logger.Write(('Suspending queue {0}' -f $Queue))
        Get-Queue $Queue | Get-Message -ResultSize Unlimited | Suspend-Message -Confirm:$false

        # Fetch suspended messages
        $logger.Write(('Fetching suspended messages from queue {0}' -f $Queue))

        $messages = @(Get-Queue $Queue | Get-Message -ResultSize Unlimited | Where-Object{$_.Status -eq "Suspended"} )

        $logger.Write( ('{0} suspended messages fetched from queue {1}' -f $messages.Count, $Queue) )

        # Export fetched messages
        $messages | ForEach-Object {$m++;Export-Message $_.Identity | AssembleMessage -Path (Join-Path -ChildPath ('{0}.eml' -f $m) -Path $Path)}
    }
    catch {
      # get error record
      [Management.Automation.ErrorRecord]$e = $_

      # retrieve information about runtime error
      $info = [PSCustomObject]@{
        Exception = $e.Exception.Message
        Reason    = $e.CategoryInfo.Reason
        Target    = $e.CategoryInfo.TargetName
        Script    = $e.InvocationInfo.ScriptName
        Line      = $e.InvocationInfo.ScriptLineNumber
        Column    = $e.InvocationInfo.OffsetInLine
      }
      
      Write-Error $info
    }
}

function Delete-Messages {
    # Delete suspended messages from queue

    $logger.Write( ('Delete  suspended messages from queue {0}' -f $Queue) )

    Get-Message -Queue $Queue -ResultSize Unlimited | Where-Object{$_.Status -eq "Suspened"} | Remove-Message -WithNDR $false -Confirm:$false
}


# MAIN ####################################################

# 1. Check export folder
Check-Folders

# 2. Check queue
if((Check-Queue -gt 0)) {
    if((Request-Choice -Caption ('Do you want to suspend and export all messages in queue {0}?' -f $Queue)) -eq 0) {
        # Yes, we want to suspend and delete messages
        Export-Messages
    }
    else {
        # No, we do not want to delete message
        $logger.Write("User choice: Do not suspend and export messages")
    }
    if($DeleteAfterExport) {
        if((Request-Choice -Caption ('Do you want to DELETE all suspended messages from queue {0}?' -f $Queue)) -eq 0) {
            $logger.Write("User choice: DELETE suspended")
            Write-Output "Suspended messages will be deleted WITHOUT sending a NDR!"
            Delete-Messages
        }
        else {
            $logger.Write("User choice: DO NOT DELETE suspended")
            Write-Output "Exported messages have NOT been deleted from queue!"
            Write-Output "Remove messages manually and be sure, if you want to send a NDR!"
        }
    }
}
else {
    Write-Output ('Queue {0} does not contain any messages' -f $Queue)
    $logger.Write(('Queue {0} does not contain any messages' -f $Queue))
}

$logger.Write("Script finished")
Write-Host "Script finished"