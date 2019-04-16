


# Allow user to select a folder to use
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
# Get the folder name and from that figure out the filenames to look for

# Change this next line to be the source of the template Assessment spreadsheet
$XLSsource = "Z:\data\customers\Microsoft\Assessment XLS\Assessment XLS v2.12.xlsm"

# Get the folder name
$folder = get-folder
write-host "Folder is ",$folder


# Use Get-ChildItem (GCI) to look in each folder and process the data
gci $folder | %{

$Path = $_.DirectoryName
$groupname = $_.BaseName
$assess_folder = $folder + '\' + $groupname
write-host "Assessment folder is ",$assess_folder

# Build the filename to save the Assessment spreadsheet based on the group name
$XLSsavepath = $assess_folder + "\" + $groupname + ".xlsm"


# Copy the template assessment spreadsheet to the new filename for the group
copy $XLSsource $XLSsavepath

$VMpath = $assess_folder + "\" + $groupname + "_vms-to-templates-mapping-csv.csv"
write-host "VM CSV path is ",$VMpath
$BYOLpath = $assess_folder + "\" + $groupname + "_byol-vms-to-templates-mapping-csv.csv"
$Storagepath = $assess_folder + "\" + $groupname + "_volume-tier-breakdown-csv.csv"
 
# Open the assessment spreadsheet
$Excel = New-Object -ComObject excel.application
$Excel.visible = $True
#Open Workbook
$workbooks = $excel.workbooks.Open($XLSsavepath)
$worksheets = $workbooks.Worksheets

#Open VMs sheet
$worksheet = $worksheets.Item("VMs")
$objRange = $worksheet.UsedRange
$numrows = $objRange.SpecialCells(11).row
$numcols = $objRange.SpecialCells(11).column

$delimiter = "," #Specify the delimiter used in the file

# Build the QueryTables.Add command and reformat the data
$TxtConnector = ("TEXT;" + $VMpath)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A15"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

# Execute & delete the import query
$query.Refresh()
$query.Delete()

 
#Open Storage sheet
$worksheet = $worksheets.Item("Storage")
$objRange = $worksheet.UsedRange
$numrows = $objRange.SpecialCells(11).row
$numcols = $objRange.SpecialCells(11).column

$delimiter = "," #Specify the delimiter used in the file

# Build the QueryTables.Add command and reformat the data
$TxtConnector = ("TEXT;" + $Storagepath)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A15"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

# Execute & delete the import query
$query.Refresh()
$query.Delete()



#Open BYOLVMs sheet
$worksheet = $worksheets.Item("BYOL VMs")
$objRange = $worksheet.UsedRange
$numrows = $objRange.SpecialCells(11).row
$numcols = $objRange.SpecialCells(11).column

$delimiter = "," #Specify the delimiter used in the file

# Build the QueryTables.Add command and reformat the data
$TxtConnector = ("TEXT;" + $BYOLpath)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A15"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

# Execute & delete the import query
$query.Refresh()
$query.Delete()


$macro = "TidyVMsTabAfterInput"
$Excel.Run($macro)
$macro = "TidyBYOLVMsTabAfterInput"
$Excel.Run($macro)
$macro = "TidyStorageTabAfterInput"
$Excel.Run($macro)



$XLSsavepath | clip
write-host "Save the spreadsheet as ",$XLSsavepath





}