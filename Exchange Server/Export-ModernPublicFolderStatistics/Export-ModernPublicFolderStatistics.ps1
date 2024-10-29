<#
    .SYNOPSIS

    This script exports statistics of modern public folders to a CSV file.

    .DESCRIPTION

    This script exports statistics of modern public folders to a CSV file. The script gathers statistics of IPM_SUBTREE and NON_IPM_SUBTREE folders.
    The script exports the following information to a CSV file:
    - Index
    - ParentPath
    - FolderName
    - ItemCount
    - TotalItemSizeInBytes
    - TotalItemSizeInMB
    - TotalDeletedItemSiteInByte
    - ContentMailboxName
    - FolderClass

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    .NOTES

    Requirements

    - Windows Server 2019+
    - Exchange Server 2016+ PowerShell Module

    Revision History
    --------------------------------------------------------------------------------
    1.0 Initial release

    .PARAMETER  ExportFile

    The name of the CSV file to which the script exports the statistics. The default value is PublicFolderStats-{0}.csv, where {0} is the current date in the format yyyy-MM-dd.

    .PARAMETER  ResultSize

    The maximum number of results to return. The default value is Unlimited.

    .EXAMPLE

    .\Export-ModernPublicFolderStatistics.ps1 -ExportFile "PublicFolderStats-2021-01-01.csv" -ResultSize 1000

    This example exports statistics of modern public folders to a CSV file named PublicFolderStats-2021-01-01.csv. The script gathers statistics of IPM_SUBTREE and NON_IPM_SUBTREE folders and returns a maximum of 1000 results.

#>

# Parameter section with examples
[CmdletBinding()]
param(
    [string]$ExportFile = ('PublicFolderStats-{0}.csv' -f (Get-Date -Format 'yyyy-MM-dd')  ),
    [string]$ResultSize = 'Unlimited' # Default value, can be changed for testing purposes
)

# Measure script running time
$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()

# some variables
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$summaryFileName = 'PublicFolderSummary.csv'

# Function that determines if to skip the given folder
function IsSkippableFolder() {
    param(
        $publicFolder
    )

    $publicFolderIdentity = $publicFolder.Identity.ToString()

    for ($index = 0; $index -lt $script:SkippedSubtree.length; $index++) {
        if ($publicFolderIdentity.StartsWith($script:SkippedSubtree[$index])) {
            return $true
        }
    }

    return $false
}

function CreateFolderObjects {

    Write-Verbose ('[{0}] Processing folder objects' -f ((Get-Date).ToString()) )
    $index = 1

    $totalFolders = $script:FolderStatistics.Count

    # Initialize progress bar
    Write-Progress -Activity "Processing Folders" -Status "Initializing..." -PercentComplete 0


    foreach ($publicFolderEntryId in $script:FolderStatistics) {

        # Update progress bar
        $percentComplete = ($index / $totalFolders) * 100
        Write-Progress -Activity "Processing Folders" -Status "Processing folder $index of $totalFolders" -PercentComplete $percentComplete

        # $publicFolderIdentity = ""
        # Write-Verbose ('[{0}] Processing [{1}]' -f ((Get-Date).ToString()), $publicFolderEntryId.Name )

        # Get public folder object
        $publicFolder = Get-PublicFolder -Identity $publicFolderEntryId.EntryId

        # Check if the folder is already processed
        $existingFolder = $script:NonIpmSubtreeFolders[$publicFolderEntryId.EntryId]

        if ($null -ne $existingFolder) {
            # NON_IPM
            $result = IsSkippableFolder($existingFolder)
            if (!$result) {
                # count
            }
        }
        else {

            # count folder based on folder class
            switch ($publicFolder.FolderClass) {
                'IPF.Appointment' { $script:ipmAppointmentCount++ }
                'IPF.StickyNote' { $script:ipmStickyNoteCount++ }
                'IPF.Contact' { $script:ipmContactCount++ }
                'IPF.Note' { $script:ipmNoteCount++ }
                '' { $script:ipmEmptyCount++ }
            }
        }

        $script:PublicFolderItemSizeInBytes += $publicFolderEntryId.TotalItemSize.ToBytes()

        # create public folder object properties
        $property = [ordered]@{
            Index                      = $index
            ParentPath                 = $publicFolder.ParentPath
            FolderName                 = $publicFolderEntryId.Name
            ItemCount                  = $publicFolderEntryId.ItemCount
            TotalItemSizeInBytes       = $publicFolderEntryId.TotalItemSize.ToBytes()
            TotalItemSizeInMB          = $publicFolderEntryId.TotalItemSize.ToMB()
            TotalDeletedItemSiteInByte = $publicFolderEntryId.TotalDeletedItemSize.ToBytes()
            ContentMailboxName         = $publicFolder.ContentMailboxName
            FolderClass                = $publicFolder.FolderClass
        }

        # create new folder object
        $newFolderObject = New-Object -TypeName PSObject -Property $property

        # add folder object to array
        $retValue = $script:ExportFolders.Add($newFolderObject)

        # increment variable
        $index++
    }

    # Complete progress bar
    Write-Progress -Activity "Processing Folders" -Status "Complete" -PercentComplete 100 -Completed

}

# Array of folder objects for exporting
$script:ExportFolders = $null

# Hash table that contains the folder list (IPM_SUBTREE via Get-PublicFolderStatistics)
$script:FolderStatistics = @{}

# Hash table that contains the folder list (NON_IPM_SUBTREE via Get-PublicFolder)
$script:NonIpmSubtreeFolders = @{}

# Hash table that contains the folder list (IPM_SUBTREE via Get-PublicFolder)
$script:IpmSubtreeFolders = @{}

# Hash table EntryId to Name to map FolderPath
$script:IdToNameMap = @{}

# summary counter
$script:ipmAppointmentCount = 0
$script:ipmStickyNoteCount = 0
$script:ipmContactCount = 0
$script:ipmEmptyCount = 0
$script:ipmNoteCount = 0
$script:PublicFolderItemSizeInBytes = 0


# Gather modern public folders of IPM_SUBTREE
Write-Progress -Activity "Fetching IPM_SUBTREE Folders" -Status "Initializing..." -PercentComplete 0
$ipmSubtreeFolderList = Get-PublicFolder "\" -Recurse -ResultSize:$ResultSize

Write-Progress -Activity "Fetching IPM_SUBTREE Folders" -Status "Id to Name mapping..." -PercentComplete 25
$ipmSubtreeFolderList | ForEach-Object { $script:IdToNameMap.Add($_.EntryId, $_.Identity.ToString()) }

Write-Verbose ('[{0}] Gathered {1} IPM_SUBTREE folders' -f ((Get-Date).ToString()), ($ipmSubtreeFolderList | Measure-Object).Count )

# Folders that are skipped while computing statistics
$script:SkippedSubtree = @("\NON_IPM_SUBTREE\OFFLINE ADDRESS BOOK", "\NON_IPM_SUBTREE\SCHEDULE+ FREE BUSY",
    "\NON_IPM_SUBTREE\schema-root", "\NON_IPM_SUBTREE\OWAScratchPad",
    "\NON_IPM_SUBTREE\StoreEvents", "\NON_IPM_SUBTREE\Events Root",
    "\NON_IPM_SUBTREE\DUMPSTER_ROOT\");

# Gather modern public folders of NON_IPM_SUBTREE
Write-Progress -Activity "Fetching NON_IPM_SUBTREE Folders" -Status "Initializing..." -PercentComplete 50
$nonIpmSubtreeFolderList = Get-PublicFolder "\NON_IPM_SUBTREE" -Recurse -ResultSize:$ResultSize

Write-Verbose ('[{0}] Gathered {1} NON_IPM_SUBTREE folders' -f ((Get-Date).ToString()), ($nonIpmSubtreeFolderList | Measure-Object).Count )

Write-Progress -Activity "Fetching NON_IPM_SUBTREE Folders" -Status "Processing..." -PercentComplete 75
foreach ($nonIpmSubtreeFolder in $nonIpmSubtreeFolderList) {
    $script:NonIpmSubtreeFolders.Add($nonIpmSubtreeFolder.EntryId, $nonIpmSubtreeFolder);
}

# Gather public folder statistics
$script:FolderStatistics = Get-PublicFolderStatistics -ResultSize:$ResultSize
Write-Verbose ('[{0}] Gathered statistics for {1} folders' -f ((Get-Date).ToString()), ($script:FolderStatistics | Measure-Object).Count )

# prepare variable
$script:ExportFolders = New-Object System.Collections.ArrayList -ArgumentList ($script:FolderStatistics.Count + 3);

# Parse folder data
CreateFolderObjects

# Export results to CSV
$exportFilePath = Join-Path -Path $ScriptDir -Childpath $ExportFile
$script:ExportFolders | Sort-Object -Property Index | Export-CSV -Path $exportFilePath -Force -NoTypeInformation -Encoding UTF8

$publicFolderItemSizeinMB =  [math]::Round($script:PublicFolderItemSizeInBytes / 1048576)

# Export summary
$property = [ordered]@{
    Date                  = (Get-Date)
    PublicFolderCount     = $ipmSubtreeFolderList.Count
    PublicFolderItemSizeInBytes = $script:PublicFolderItemSizeInBytes
    PublicFolderItemSizeInMB = $publicFolderItemSizeinMB
   # PublicFolderItemSizeInGB = [INT](($script:PublicFolderItemSizeInBytes).ToMB()/1024)
    DefaultFolderCount    = $script:ipmNoteCount
    CalendarFolderCount   = $script:ipmAppointmentCount
    ContactFolderCount    = $script:ipmContactCount
    StickyNoteFolderCount = $script:ipmStickyNoteCount
    IPMEmptyFolderCount   = $script:ipmEmptyCount
}

# Export summary object to CSV and append to existing file
$summaryObject = New-Object -TypeName PSObject -Property $property
$summaryObject | Export-CSV -Path (Join-Path -Path $ScriptDir -ChildPath $summaryFileName) -Append -Force -NoTypeInformation -Encoding UTF8

# Finish
$StopWatch.Stop()
Write-Verbose -Message ('It took {0:00}:{1:00}:{2:00} to run the script.' -f $StopWatch.Elapsed.Hours, $StopWatch.Elapsed.Minutes, $StopWatch.Elapsed.Seconds)