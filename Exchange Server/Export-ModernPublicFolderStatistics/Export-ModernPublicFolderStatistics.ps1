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
    1.1 Logging added

    .PARAMETER  ExportFile

    The name of the CSV file to which the script exports the statistics. The default value is PublicFolderStats-{0}.csv, where {0} is the current date in the format yyyy-MM-dd.
    The script copies the file to the archive folder.

    .PARAMETER ExportFileDefault

    The name of the CSV file containing the most current export for further processing by other scripts.

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
    [string]$ExportFilenameCurrent = 'PublicFolderStats-Current.csv',
    [string]$ResultSize = 'Unlimited' # Default value, can be changed for testing purposes
)

# Measure script running time
$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()

# some variables
$scriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$summaryFileName = 'PublicFolderSummary.csv'

function Import-RequiredModules {

  # Import central logging functions
  if($null -ne (Get-Module -Name GlobalFunctions -ListAvailable).Version) {
    Import-Module -Name GlobalFunctions
  }
  else {
    Write-Warning -Message 'Unable to load GlobalFunctions PowerShell module.'
    Write-Warning -Message 'Please check http://bit.ly/GlobalFunctions for further instructions'
    exit
  }

}


# Function that determines if to skip the given folder
function IsSkippableFolder() {
    [CmdletBinding()]
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

function Copy-ToArchive {
    [CmdletBinding()]
    param(
        [string]$SourceFilename,
        [string]$archiveFolderName = 'Archive'
    )

    if(-not (Test-Path -Path (Join-Path -Path $scriptDir -ChildPath $archiveFolderName)  ) ) {
        Write-Verbose ('[{0}] Creating archive folder' -f ((Get-Date).ToString()))
        New-Item -Path $scriptDir -Name $archiveFolderName -ItemType "directory" -Confirm:$false
    }

    try{
        if (Test-Path -Path (Join-Path -Path (Join-Path -Path $scriptDir -ChildPath $archiveFolderName ) -ChildPath $ExportFilenameCurrent) ) {
            Remove-Item -Path (Join-Path -Path (Join-Path -Path $scriptDir -ChildPath $archiveFolderName ) -ChildPath $ExportFilenameCurrent) -Force -Confim:$false
        }
        $null = Copy-Item $SourceFilename -Destination (Join-Path -Path $scriptDir -ChildPath $archiveFolderName) -Force -Confirm:$false
    }
    catch{
        Write-Error ('The script cannot copy the CSV file to {0}' -f (Join-Path -Path $scriptDir -ChildPath $archiveFolderName) )
    }

    try {
        Write-Verbose ('Rename: {0}' -f (Join-Path -Path $scriptDir -ChildPath $ExportFilenameCurrent)  )
        $null = Remove-Item -Path (Join-Path -Path $scriptDir -ChildPath $ExportFilenameCurrent) -Force -Confirm:$false #-ErrorAction SilentlyContinue

        $null = Rename-Item -Path $SourceFilename -NewName $ExportFilenameCurrent -Force -Confirm:$false
    }
    catch{
        Write-Error ('The script cannot rename the CSV file to {0}' -f (Join-Path -Path $scriptDir -ChildPath $ExportFilenameCurrent) )
    }

}

### MAIN ###########################
Import-RequiredModules

# Create new logger
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
# Purge logs depening on LogFileRetention
$logger.Purge()
$logger.Write('Script started')

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
Write-Progress -Activity "Gathering Public Folders" -Status "IPM_SUBTREE folder... [be patient]" -PercentComplete 25
$ipmSubtreeFolderList = Get-PublicFolder "\" -Recurse -ResultSize:$ResultSize
$ipmSubtreeFolderList | ForEach-Object { $script:IdToNameMap.Add($_.EntryId, $_.Identity.ToString()) }

$logger.Write( ('Gathered {0} IPM_SUBTREE folders' -f ($ipmSubtreeFolderList | Measure-Object).Count  ) )
Write-Verbose ('[{0}] Gathered {1} IPM_SUBTREE folders' -f ((Get-Date).ToString()), ($ipmSubtreeFolderList | Measure-Object).Count )

# Folders that are skipped while computing statistics
$script:SkippedSubtree = @("\NON_IPM_SUBTREE\OFFLINE ADDRESS BOOK", "\NON_IPM_SUBTREE\SCHEDULE+ FREE BUSY",
    "\NON_IPM_SUBTREE\schema-root", "\NON_IPM_SUBTREE\OWAScratchPad",
    "\NON_IPM_SUBTREE\StoreEvents", "\NON_IPM_SUBTREE\Events Root",
    "\NON_IPM_SUBTREE\DUMPSTER_ROOT\");

# Gather modern public folders of NON_IPM_SUBTREE
Write-Progress -Activity "Gathering Public Folders" -Status "NON_IPM_SUBTREE folder... [be patient]" -PercentComplete 50
$nonIpmSubtreeFolderList = Get-PublicFolder "\NON_IPM_SUBTREE" -Recurse -ResultSize:$ResultSize
$logger.Write( ('Gathered {0} NON_IPM_SUBTREE folders' -f ($nonIpmSubtreeFolderList | Measure-Object).Count  ) )
Write-Verbose ('[{0}] Gathered {1} NON_IMP_SUBTREE folders' -f ((Get-Date).ToString()), ($nonIpmSubtreeFolderList | Measure-Object).Count )

foreach ($nonIpmSubtreeFolder in $nonIpmSubtreeFolderList) {
    $script:NonIpmSubtreeFolders.Add($nonIpmSubtreeFolder.EntryId, $nonIpmSubtreeFolder);
}

# Gather public folder statistics
Write-Progress -Activity "Gathering Public Folders" -Status "Public Folder Statistics... [be patient]" -PercentComplete 75
$logger.Write( 'Gathering Public Folder Statistics' )

$script:FolderStatistics = Get-PublicFolderStatistics -ResultSize:$ResultSize

$logger.Write( ('Gathered statistics for {0} folders' -f ($script:FolderStatistics | Measure-Object).Count) )
Write-Verbose ('[{0}] Gathered statistics for {1} folders' -f ((Get-Date).ToString()), ($script:FolderStatistics | Measure-Object).Count )

# prepare variable
$script:ExportFolders = New-Object System.Collections.ArrayList -ArgumentList ($script:FolderStatistics.Count + 3)

# Parse folder data
CreateFolderObjects

# Export results to CSV
$exportFilePath = Join-Path -Path $scriptDir -Childpath $ExportFile
$script:ExportFolders | Sort-Object -Property Index | Export-CSV -Path $exportFilePath -Force -NoTypeInformation -Encoding UTF8

Copy-ToArchive -SourceFilename $exportFilePath

# Fetch Root folders
$rootFolderCount = (Get-PublicFolder \ -GetChildren -ResultSize Unlimited).Count

# Do some calculations
$publicFolderItemSizeinMB = [math]::Round($script:PublicFolderItemSizeInBytes / 1048576)
$publicFolderItemSizeinGB = [math]::Round($publicFolderItemSizeinMB / 1024)

# Export summary
$property = [ordered]@{
    Date                  = (Get-Date -Format 'dd.MM.yyyy')
    Count     = $ipmSubtreeFolderList.Count
    SizeInBytes = $script:PublicFolderItemSizeInBytes
    SizeInMB = $publicFolderItemSizeinMB
    SizeInGB = $publicFolderItemSizeinGB
    DefaultFolders    = $script:ipmNoteCount
    CalendarFolders   = $script:ipmAppointmentCount
    ContactFolders    = $script:ipmContactCount
    StickyNoteFolders = $script:ipmStickyNoteCount
    NoFolderType   = $script:ipmEmptyCount
    RootFolders = $rootFolderCount
}

# Export summary object to CSV and append to existing file
$summaryObject = New-Object -TypeName PSObject -Property $property
$summaryObject | Export-CSV -Path (Join-Path -Path $scriptDir -ChildPath $summaryFileName) -Append -Force -NoTypeInformation -Encoding UTF8

# Finish
$StopWatch.Stop()
$logger.Write( ('Script runtime {0:00}:{1:00}:{2:00}' -f $StopWatch.Elapsed.Hours, $StopWatch.Elapsed.Minutes, $StopWatch.Elapsed.Seconds) )
Write-Verbose -Message ('It took {0:00}:{1:00}:{2:00} to run the script.' -f $StopWatch.Elapsed.Hours, $StopWatch.Elapsed.Minutes, $StopWatch.Elapsed.Seconds)