<#
.SYNOPSIS
    Creates and emails a WSUS report formatted in Excel
.DESCRIPTION
    The script will generate a list of Servers managed by WSUS, create and email a report formatted in Excel without installing Excel.
    Report lists patch status of each server along with a few pivot charts to choose from.
.EXAMPLE
    Run this PS Script on WSUS Server
.INPUTS
    Read from WSUS Server
.OUTPUTS
    Excel report saved to C:\Reports folder
.NOTES
    9/22/2021 Mark Tellier

    - Script must be run as Administrator
    - Requires PoshWSUS and ImportExcel modules from PowerShell Gallery
    - PoshWSUS requires Administrator mode
    - Modify Variables as required
#>

# --- LOAD REQUIRED MODULES
If ( ! (Get-module PoshWSUS )) {
    Import-Module PoshWSUS
}
If ( ! (Get-module ImportExcel )) {
    Import-Module ImportExcel
}

# --- MAIN VARIABLES
$today = get-date -Format "MM-dd-yyyy"
$reportTitle = "Security Update Report - ACME Production Infrastructure $today"
$tableName = "UpdateReport"
$worksheetName = "Servers"
$report = @()
$lineNo = 0
$WSUShost = $env:COMPUTERNAME.ToUpper()
$xlsxPath = "C:\Reports\" + $WSUShost + "_" + $today + ".xlsx"

# --- EMAIL VARIABLES
$smtpServer = "10.10.10.10"
$smtpFrom = "noreploy@acme.com"
$smtpTo = 'The Dude <thedude@acme.com>', 'Other Dude <otherdude@acme.com>'
$smtpSubject = "WSUS Report"
$l1 = "This is an automated report generated $today from patching server $WSUShost.`n`n"
$l2 = "After opening the attachment, you will need to 'Enable Editing' when prompted in order to view the chart.`n`n"
$l3 = "Instructions for printing:`n"
$l4 = "1. Save attached Excel spreadsheet to local disk.`n"
$l5 = "2. Select 'Page Layout'.`n"
$l6 = "3. Set Orientation > 'Landscape`n"
$l7 = "4. Highlight entire table, including headers, select 'Print Area' > 'Set Print Area'.`n"
$l8 = "5. Select 'Print Titles' and highlight first two rows.`n"
$l9 = "6. Click 'Header/Footer', under footer > 'Confidential (predefined)', select 'Print Preview'`n"
$l10 = "7. Select PDF printer, 'LandscapeOrientation', Scaling 'Fit all Columns on One Page', Click 'Print'.`n"
$smtpBody = $l1 + $l2 + $l3 + $l4 + $l5 + $l6 + $l7 + $l8 + $l9 + $l10

# --- READ INFO FROM WSUS
Connect-PSWSUSServer -WsusServer $WSUShost -Port 8530
$clientList = Get-PSWSUSUpdateSummaryPerClient | Sort-Object Computer

# --- FORMAT DATA
foreach ( $client in $clientList ) {

    $extInfo = Get-PSWSUSClient -ComputerName $client.Computer
    $lineNo++

    $items = [ordered]@{

        "Index"             = $lineNo
        "Computer"          = $client.Computer.tolower()
        "IP Address"        = $extInfo.IPAddress
        "Operating System"  = $extInfo.OSDescription
        "Last Sync"         = $extInfo.LastSyncTime.tolocaltime()
        "Last Result"       = $extInfo.LastSyncResult
        "Needed"            = $client.NeededCount
        "Downloaded"        = $client.DownloadedCount
        "Failed"            = $client.FailedCount
        "Installed"         = $client.InstalledCount 
        "Pending Reboot"    = $client.PendingReboot
        
    }

    $report += New-Object -TypeName psobject -Property $items

}

# --- DELETE TARGET FILE IF EXISTS
Write-Verbose -Verbose -Message "Save location: $xlsxPath"
Remove-Item $xlsxPath -ErrorAction Ignore

# --- CONVERT DATA TO XLSX
$excelParam = @{
    Path                = $xlsxPath
    WorksheetName       = $worksheetName
    TableName           = $tableName
    Title               = $reportTitle
}

$excel = $report | Export-Excel -PassThru -AutoSize -DisplayPropertySet @excelParam

# --- FORMAT COLUMNS
Set-ExcelColumn -ExcelPackage $excel -WorksheetName $worksheetName -Column 1 -HorizontalAlignment Center
Set-ExcelColumn -ExcelPackage $excel -WorksheetName $worksheetName -Column 5 -HorizontalAlignment Left -Width 19

$count = 7
while ($count -le 11) {
    Set-ExcelColumn -ExcelPackage $excel -WorksheetName $worksheetName -Column $count -HorizontalAlignment Center
    $count++
}

# --- CENTER TITLE ROW
$sheet1 = $excel.Workbook.Worksheets["Servers"]
Set-ExcelRange -Address $sheet1.Cells["A1:K1"] -Merge -HorizontalAlignment Center -FontSize 18

# --- DEFINE PIVOT TABLE PARAMETERS
$pivotOne = @{
    ExcelPackage        = $excel
    PivotTableName      = "P1"
    SourceRange         = $excel.Workbook.Worksheets["Servers"].Tables["UpdateReport"]
    PivotRows           = "Operating System"
    PivotData           = @{'Operating System' = 'count'}
    ChartType           = "BarClustered3D"
    PivotTableStyle     = "Dark2"
    ChartTitle          = "Windows OS"
}

$pivotTwo = @{
    ExcelPackage        = $excel
    PivotTableName      = "P2"
    SourceRange         = $excel.Workbook.Worksheets["Servers"].Tables["UpdateReport"]
    PivotRows           = "Operating System"
    PivotData           = @{'Operating System' = 'count'}
    ChartType           = "Doughnut"
    PivotTableStyle     = "Medium17"
    ChartTitle          = "OS Versions"
}

$pivotThree = @{
    ExcelPackage        = $excel
    PivotTableName      = "P3"
    SourceRange         = $excel.Workbook.Worksheets["Servers"].Tables["UpdateReport"]
    PivotRows           = "Operating System"
    PivotData           = @{'Operating System' = 'count'}
    ChartType           = "Pie"
    PivotTableStyle     = "Medium3"
    ChartTitle          = "Windows Versions"
}

$pivotFour = @{
    ExcelPackage        = $excel
    PivotTableName      = "P4"
    SourceRange         = $excel.Workbook.Worksheets["Servers"].Tables["UpdateReport"]
    PivotRows           = "Operating System"
    PivotData           = @{'Operating System' = 'count'}
    ChartType           = "Pie3D"
    PivotTableStyle     = "Dark11"
    ChartTitle          = "Operating Systems"
}

# --- DEFINE PIVOT TABLES
Add-PivotTable -IncludePivotChart -Activate  @pivotOne
Add-PivotTable -IncludePivotChart -Activate  @pivotTwo
Add-PivotTable -IncludePivotChart -Activate  @pivotThree -ShowPercent
Add-PivotTable -IncludePivotChart -Activate  @pivotFour -ShowPercent

# --- SELECT FIRST SHEET AS ACTIVE TAB
Select-Worksheet -ExcelPackage $excel -WorksheetName "Servers"

# --- SAVE CHANGES, CLEAR MEMORY
Close-ExcelPackage $excel

# --- EMAIL REPORT
$smtpParam = @{
    SmtpServer      = $smtpServer
    From            = $smtpFrom
    To              = $smtpTo
    Subject         = $smtpSubject
    Body            = $smtpBody
    Attachments     = $xlsxPath
}

Send-MailMessage @smtpParam
