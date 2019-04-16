# Fix_Storage.ps1
# Fixes cloud pricing in Cloud Migration Assessment spreadsheets
# Richard Stinton - April 2019
#

Function Get-FileName($initialDirectory)
{
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Please Select File"
    $OpenFileDialog.InitialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    # Out-Null supresses the "OK" after selecting the file.
    $OpenFileDialog.ShowDialog() | Out-Null
    $Global:SelectedFile = $OpenFileDialog.FileName
}
# Reads the XLSM file used for Turbo Cloud Migration assessments
#
#
# Edit here whether to use hard-coded input file or file browser

Get-FileName("Z:\data\customers\Microsoft\pstest\Carli")
# write-host $Global:SelectedFile
#
#
#Location of the Excel file I want to edit
#$FileLoc = "Z:\data\customers\microsoft\pstest\test.xlsm"
$FileLoc = $Global:SelectedFile
#
#
#Create Excel Com Object, and display it
$excel = new-object -com Excel.Application
$excel.visible = $true
$FileLoc

#Open Workbook,
$workbooks = $excel.workbooks.Open($FileLoc)
$worksheets = $workbooks.Worksheets



#Open Storage sheet
$worksheet = $worksheets.Item("Storage")
$objRange = $worksheet.UsedRange
$numrows = $objRange.SpecialCells(11).row
$numcols = $objRange.SpecialCells(11).column
write-host "Processing ",$numrows," rows of data for Storage sheet"
#Select the range we want to look at
#In this example, I am only checking within one Column
$range = $worksheet.Range("A1","M"+$numrows)
# $range.cells.item(1,1).text
# Cycle through rows reading the VMname, Volume Tier, Size, Region, Cost
foreach ($row in $range.rows){
    #$row.Row
    $Volume=$row.Cells.item(1)
    $AllTier=$row.Cells.item(5)
    $AllSize=$row.Cells.item(6)
    $AllReg=$row.Cells.item(7)
    $AllPrice=$row.Cells.item(8)
    $ConsTier=$row.Cells.item(9)
    $ConsSize=$row.Cells.item(10)
    $ConsReg=$row.Cells.item(11)
    $ConsPrice=$row.Cells.item(12)
    

# Set Allocation tier to Managed_Premium and update price
# $
$mystring = $AllSize.value2
IF($row.row -gt 7) {
$worksheet.Cells.Item($row.Row,5)="MANAGED_PREMIUM"
IF($mystring -le 32767) { 
    $worksheet.Cells.Item($row.Row,8)="5.807"
} elseif($mystring -le 65536) {
    $worksheet.Cells.Item($row.Row,8)="11.227"
} elseif($mystring -le 131072) {
    $worksheet.Cells.Item($row.Row,8)="21.680"
} elseif($mystring -le 250880) {
    $worksheet.Cells.Item($row.Row,8)="41.811"
} elseif($mystring -le 524288) {
    $worksheet.Cells.Item($row.Row,8)="80.540"
} elseif($mystring -le 1048576) {
    $worksheet.Cells.Item($row.Row,8)="148.680"      
} elseif($mystring -le 2097152) {
    $worksheet.Cells.Item($row.Row,8)="284.937"
} elseif($mystring -le 4190208) {
    $worksheet.Cells.Item($row.Row,8)="545.097" 
}
    }
        
}

