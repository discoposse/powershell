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

$filepath = get-folder
$groupname = $filepath.Split('\')[$($filepath.Split('\').Count-1)]
$filepath
$groupname


#$filepath = "Z:\data\customers\Microsoft\Proximus\18-3-19\plans\Carli\OSSINLIN-01\"
#$groupname = "VMs_NC1 Carli_NC1-SRV-OSSINLIN-01"

$XLSsavepath = $filepath + "\" + $groupname + ".xlsm"
$XLSsource = "Z:\data\customers\Microsoft\Assessment XLS\Assessment XLS v2.12.xlsm"

copy $XLSsource $XLSsavepath
$VMpath = $filepath + "\" + $groupname + "_vms-to-templates-mapping-csv.csv"
$BYOLpath = $filepath + "\" + $groupname + "_byol-vms-to-templates-mapping-csv.csv"
$Storagepath = $filepath + "\" + $groupname + "_volume-tier-breakdown-csv.csv"
 
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


$savepath = $groupname + '.xlsm'
$savepath | clip
write-host "Save the spreadsheet as ",$savepath

