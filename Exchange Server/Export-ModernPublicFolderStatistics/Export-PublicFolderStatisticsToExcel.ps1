# Parameter section
[CmdletBinding()]
param(
    [switch]$OpenExcel,
    [switch]$SendMail,
    [string]$MailFrom = '',
    [string]$MailTo = '',
    [string]$MailServer = ''
)

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

# Use current date as file timestamp
$now = Get-Date -Format 'yyyy-MM-dd'

$cssFile = Join-Path -Path $ScriptDir -ChildPath styles.css
$reportTitle = "Public Folder Statistics - $($now)"

function Import-PublicFolderStatsToExcelWithChart {
    param (
        [string]$csvPath = "PublicFolderSummary.csv",
        [string]$excelPath = "PublicFolderSummary.xlsx"
    )

    # Check if the ImportExcel module is installed
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "ImportExcel module is not installed. Installing now..."
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
    }

    Remove-Item $excelPath -ErrorAction Ignore

    # Import data from CSV
    $data = Import-Csv -Path $csvPath

    $workSheetName = "PublicFolderStats"

    # Export data to Excel
    $excel = $data | Export-Excel -Path $excelPath -WorksheetName $workSheetName -ClearSheet -TableName PublicFolder -AutoSize -PassThru -AutoNameRange



    # some chart parameters
    $width = 1000
    $height = 500
    $row = 1
    $column = 1

    # Add chart sheet
    $null = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Overview' -Activate -MoveToEnd

    # Add a line chart to the Excel file
    $excelChartParams = @{
        Worksheet    = $excel.Overview
        ChartType    = "LineMarkersStacked"
        Title        = "Entwicklung der Public Folder (Anzahl)"
        #XRange       = "A2:A" + ($data.Count + 1)
        #YRange       = "B2:B" + ($data.Count + 1) #+ ",D2:D" + ($data.Count + 1)  # PublicFolderCount and PublicFolderItemSizeInMB columns
        XRange    = ('{0}!Date' -f $workSheetName)
        YRange    = ('{0}!Count' -f $workSheetName)# ,('{0}!SizeInMB' -f $workSheetName)
        SeriesHeader = "Anzahl Public Folder"# , "Größe in MB"
        XAxisTitleText = "Datum"
        XAxisTitleSize = 10
        #XAxisNumberformat = "Date-Time"
        YAxisTitleText = "Anzahl"
        YAxisTitleSize = 10
        LegendPosition = "bottom"
        Row          = $row
        Column       = $column
        Width        = $width
        Height       = $height

    }

    Add-ExcelChart @excelChartParams

    # Add chart sheet
    $null = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Size' -Activate -MoveToEnd

    # Add a line chart to the Excel file
    $excelChartParams = @{
        Worksheet    = $excel.Size
        ChartType    = "LineMarkersStacked"
        Title        = "Entwicklung der Public Folder (Size)"
        #XRange       = "A2:A" + ($data.Count + 1)
        #YRange       = "B2:B" + ($data.Count + 1) #+ ",D2:D" + ($data.Count + 1)  # PublicFolderCount and PublicFolderItemSizeInMB columns
        XRange    = ('{0}!Date' -f $workSheetName)
        YRange    = ('{0}!SizeInMB' -f $workSheetName)
        SeriesHeader = "Size in MB"
        XAxisTitleText = "Datum"
        XAxisTitleSize = 10
        YAxisNumberformat = "#,##0"
        YAxisTitleText = "Size [MB]"
        YAxisTitleSize = 10
        LegendPosition = "bottom"
        Row          = $row
        Column       = $column
        Width        = $width
        Height       = $height

    }

    Add-ExcelChart @excelChartParams

    # Add chart sheet
    $null = Add-Worksheet -ExcelPackage $excel -WorksheetName 'First Level' -Activate -MoveToEnd

    # Add a line chart to the Excel file
    $excelChartParams = @{
        Worksheet    = $excel."First Level"
        ChartType    = "LineMarkersStacked"
        Title        = "Public Folder Erste Ebene (Anzahl)"
        #XRange       = "A2:A" + ($data.Count + 1)
        #YRange       = "B2:B" + ($data.Count + 1) #+ ",D2:D" + ($data.Count + 1)  # PublicFolderCount and PublicFolderItemSizeInMB columns
        XRange    = ('{0}!Date' -f $workSheetName)
        YRange    = ('{0}!RootFolders' -f $workSheetName)
        SeriesHeader = "Anzahl"
        XAxisTitleText = "Datum"
        XAxisTitleSize = 10
        YAxisNumberformat = "#,##0"
        YAxisTitleText = "Anzahl"
        YAxisTitleSize = 10
        LegendPosition = "bottom"
        Row          = $row
        Column       = $column
        Width        = $width
        Height       = $height

    }

    Add-ExcelChart @excelChartParams

    if($OpenExcel) {
        # Open the Excel file
        Close-ExcelPackage $excel -Show
    }
    esle{
        # Save the Excel file
        Close-ExcelPackage $excel
    }

    Write-Host "Data and chart have been successfully exported to $excelPath"
}

# Call the function to create the Excel file
Import-PublicFolderStatsToExcelWithChart

if ($SendMail) {



    $head = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($reportTitle)</title>
<style type="text/css">$(Get-Content $cssFile)</style>
<body><h1 align=""center"">$($reportTitle)</h1>
<p>Here are the current public folder statistics.</p>
"@

    [string]$htmlreport = ConvertTo-Html -Body $html -Head $head -Title $reportTitle

    Send-Mail -From $MailFrom -To $MailTo -SmtpServer $MailServer -MessageBody $htmlreport -Subject $reportTitle -Attachments $excelPath -BodyAsHtml
}