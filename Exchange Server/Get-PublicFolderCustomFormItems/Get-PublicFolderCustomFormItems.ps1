<#
.SYNOPSIS
    Get public folder items of a specific type and export the results to a CSV file.

.DESCRIPTION
This script retrieves all public folders under a specified path, counts the items of a specific type in each folder, and exports the results to a CSV file.
It is particularly useful for counting custom forms in public folders in an Exchange environment.
You must have the Exchange Management Shell or the Exchange Online PowerShell module installed to run this script.

.LINK
https://granikos.eu/go/ZtsJ

.PARAMETER PublicFolderPath
The path to the public folder from which to retrieve items.

.PARAMETER ItemType
The type of item to count in the public folders. Default is 'IPM.Post.FORMNAME'.
Adjust this parameter to specify the item type you are interested in.
You can find the form names by using the `Get-PublicFolderItemStatistics` cmdlet.

.EXAMPLE
.\Get-PublicFolderCustomFormItems.ps1 -PublicFolderPath '\Department\HR' -ItemType 'IPM.Post.FORMNAME'

Retrieves all public folders under '\Department\HR', counts the items of type 'IPM.Post.FORMNAME' in each folder, and exports the results to a CSV file. Adjust the name of the item type as needed.
#>

[CmdletBinding()]
param(
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Path to public folder")]
    [ValidateNotNull()]
    [string]$PublicFolderPath,
    [string]$ItemType = 'IPM.Post.FORMNAME'
)

$i = 1
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$publicFolders = Get-PublicFolder $PublicFolderPath -Recurse -ResultSize Unlimited

# Ensure we have some public folders to work with
$publicFolderCount = ($publicFolders | Measure-Object).Count

Write-Host ('Fetched {0} public folders' -f $publicFolderCount )

# Initialize an array to hold the folder information
$objFolderItems = @()

if($publicFolderCount -gt 0){

    foreach($folder in $publicFolders) {

        Write-Progress -Activity ('Working on ({2}/{1}): {0}' -f $folder.Identity, $publicFolderCount, $i ) -PercentComplete (($i / $publicFolderCount) * 100)

        $folderItems  = Get-PublicFolderItemStatistics -Identity $folder.Identity

        $folderItemTypeCount = ($folderItems | Group-Object ItemType | Where-Object{$_.name -eq $ItemType}).Count

        $objFolderItems += [PSCustomObject]@{
            FolderPath = $folder.Identity
            ItemCount = ($FolderItems |Measure-Object).Count
            ItemTypeCount =$folderItemTypeCount
            ItemCreationTime = ($folderItems | ?{$_.ItemType -eq $ItemType} | Sort-Object CreationTime -Descending | Select-Object -First 1).CreationTime
            ItemLastModificationTime = ($folderItems | ?{$_.ItemType -eq $ItemType} | Sort-Object LastModificationTime -Descending | Select-Object -First 1).LastModificationTime
        }

        $i++
    }

    # Output the results to the console
    $objFolderItems | Format-Table -autosize

    # Export the results to a CSV file
    $objFolderItems | export-csv -Path (Join-Path -Path $scriptDir -ChildPath 'CustomFormsReport.csv') -Encoding UTF8 -NoTypeInformation -Force

}
else {
    Write-Host 'No public folders found.'
}