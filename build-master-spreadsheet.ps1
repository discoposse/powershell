# Script to cycle through a folder structure of assessment spreadsheets
# and build a Master spreadsheet of results and a graph
# Richard Stinton - April 2019
#
# Usage:
# 1. Asks to select the top level folder of the assessment folder structure
# 2. Copies the template assessment XLSM over to that folder from a hard-coded location (see below at line 36)
# 3. Uses the Get-ChildItem to cycle through each sub-folder looking for the assessment spreadsheet which it opens
#    and reads to get the totals data
    
    
    
# Get the top level folder to work on

    Function Get-Folder($initialDirectory)

{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}


# Start of main script

# Get the name of the folder to work on
$folder_root = get-folder

# Update the next line with the location of your master assessment XLS
$masterxls = "Z:\data\customers\Microsoft\Assessment XLS\Master Assessment v1.0.xlsx" 
$newmasterxls = $folder_root +'\'+'master.xlsx'
copy $masterxls $newmasterxls
$Excel1 = New-Object -ComObject excel.application
$Excel1.visible = $True
#Open Master Assessment Workbook
$workbooks1 = $excel1.workbooks.Open($newmasterxls)
$worksheets1 = $workbooks1.Worksheets
$worksheet1 = $worksheets1.Item("Sheet1")


$folders = Get-ChildItem $folder_root -Exclude '*.xlsx' | sort

$xlsrow = 15

foreach ($folder in $folders) {

$xlsname = $folder.FullName + "\" + $folder.BaseName + ".xlsm"


$Excel = New-Object -ComObject excel.application
$Excel.visible = $False
#Open Workbook
$workbooks = $excel.workbooks.Open($XLSname)
$worksheets = $workbooks.Worksheets


#Open Overall Totals sheet
$worksheet = $worksheets.Item("Overall Totals")
$objRange = $worksheet.UsedRange
$numrows = $objRange.SpecialCells(11).row
$numrows
$numcols = $objRange.SpecialCells(11).column
$numcols

$range = $worksheet.Range("A1","M"+$numrows)
$LS = $range.cells.item(7,5).text
$RS = $range.cells.item(8,5).text
$AHUB = $range.cells.item(9,5).text
$RI = $range.cells.item(10,5).text
$PO = $range.cells.item(11,5).text
$ODP = $range.cells.item(12,5).text

$worksheet = $worksheets.Item("VMs")
$objRange = $worksheet.UsedRange
$numrows = $objRange.SpecialCells(11).row
$numrows
$numcols = $objRange.SpecialCells(11).column
$numcols

$worksheet1.Cells.Item($xlsrow,1)=$folder.BaseName
$worksheet1.Cells.Item($xlsrow,3)=$LS
$worksheet1.Cells.Item($xlsrow,4)=$RS
$worksheet1.Cells.Item($xlsrow,5)=$AHUB
$worksheet1.Cells.Item($xlsrow,6)=$RI
# $worksheet1.Cells.Item($xlsrow,7)=$PO
# $worksheet1.Cells.Item($xlsrow,8)=$ODP
# $worksheet1.Cells.Item($xlsrow,9)=$numrows-7
$worksheet1.Cells.Item($xlsrow,7)=$numrows-7

$xlsrow++

$Excel.Workbooks.Close()
$Excel.Quit()

}
