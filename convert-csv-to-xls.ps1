
gci "Z:\data\customers\Microsoft\pstest\csvs" | %{

$Path = $_.DirectoryName
$filename = $_.BaseName
write-host "Path", $Path
write-host "Filename", $filename

#Define locations and delimiter
$csv = $_.FullName #Location of the source file
$csv
$xlsx = "$Path/$filename.xlsx" # Names & saves Excel file same name/location as CSV
#$xlsx = "c:/path/to/save/files/$filename.xlsx" # Names Excel file same name as CSV

$delimiter = "," #Specify the delimiter used in the file

# Create a new Excel workbook with one empty sheet
$excel = New-Object -ComObject excel.application
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)

# Build the QueryTables.Add command and reformat the data
$TxtConnector = ("TEXT;" + $csv)
$TxtConnector 
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A15"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

# Execute & delete the import query
$query.Refresh()
$query.Delete()

# Save & close the Workbook as XLSX.
$Workbook.SaveAs($xlsx,51)
$excel.Quit()

}
