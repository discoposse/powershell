# JiggeryPockery.ps1
# Fixes empty cell issues in Cloud Migration Assessment spreadsheets
# Richard Stinton - March 2019
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

Get-FileName("Z:\data\customers\microsoft")
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

#Open Workbook,
$workbooks = $excel.workbooks.Open($FileLoc)
$worksheets = $workbooks.Worksheets
#Open VMs sheet
$worksheet = $worksheets.Item("VMs")
$objRange = $worksheet.UsedRange
$numrows = $objRange.SpecialCells(11).row
$numcols = $objRange.SpecialCells(11).column
write-host "Processing ",$numrows," rows of data for VMs sheet"
#Select the range we want to look at

$range = $worksheet.Range("A1","M"+$numrows)
# $range.cells.item(1,1).text
# Cycle through rows reading the VMname, Allocation data and Consumption
foreach ($row in $range.rows){
    #$row.Row
    $VMname=$row.Cells.item(1)
    $AllInst=$row.Cells.item(4)
    $AllReg=$row.Cells.item(5)
    $AllPrice=$row.Cells.item(6)
    $ConsInst=$row.Cells.item(7)
    $ConsReg=$row.Cells.item(9)
    $ConsPrice=$row.Cells.item(10)
   

# Check if Allocation instance name is blank
# If it is replace the Allocation values with the Consumption values
$mystring = $AllInst.text
IF([string]::IsNullOrEmpty($mystring)) {            
    Write-Host "Empty Allocation cell found for server ",$VMname.text," at row ",$row.Row
    $worksheet.Cells.Item($row.Row,4)=$ConsInst.Text
    $worksheet.Cells.Item($row.Row,5)=$ConsReg.Text
    $worksheet.Cells.Item($row.Row,6)=$ConsPrice.Text
} 

# Check if Consumption instance name is blank
# If it is replace the Consumption values with the Allocation values
$mystring = $ConsInst.text
IF([string]::IsNullOrEmpty($mystring)) {            
    Write-Host "Empty Consumption cell found for server ",$VMname.text," at row ",$row.Row
    $worksheet.Cells.Item($row.Row,7)=$AllInst.Text
    $worksheet.Cells.Item($row.Row,9)=$AllReg.Text
    $worksheet.Cells.Item($row.Row,10)=$AllPrice.Text
} 
        }


#Open BYOL VMs sheet
$worksheet = $worksheets.Item("BYOL VMs")
$objRange = $worksheet.UsedRange
$numrows = $objRange.SpecialCells(11).row
$numcols = $objRange.SpecialCells(11).column
write-host "Processing ",$numrows," rows of data for BYOL VMs sheet"
#Select the range we want to look at
#In this example, I am only checking within one Column
$range = $worksheet.Range("A1","M"+$numrows)
# $range.cells.item(1,1).text
# Cycle through rows reading the VMname, Allocation data and Consumption
foreach ($row in $range.rows){
    #$row.Row
    $VMname=$row.Cells.item(1)
    $AllInst=$row.Cells.item(4)
    $AllReg=$row.Cells.item(5)
    $AllPrice=$row.Cells.item(6)
    $ConsInst=$row.Cells.item(7)
    $ConsReg=$row.Cells.item(9)
    $ConsPrice=$row.Cells.item(10)
    


# Check if Allocation instance name is blank
# If it is replace the Allocation values with the Consumption values
$mystring = $AllInst.text
IF([string]::IsNullOrEmpty($mystring)) {            
    Write-Host "Empty Allocation cell found for server ",$VMname.text," at row ",$row.Row
    $worksheet.Cells.Item($row.Row,4)=$ConsInst.Text
    $worksheet.Cells.Item($row.Row,5)=$ConsReg.Text
    $worksheet.Cells.Item($row.Row,6)=$ConsPrice.Text
} 

# Check if Consumption instance name is blank
# If it is replace the Consumption values with the Allocation values
$mystring = $ConsInst.text
IF([string]::IsNullOrEmpty($mystring)) {            
    Write-Host "Empty Consumption cell found for server ",$VMname.text," at row ",$row.Row
    $worksheet.Cells.Item($row.Row,7)=$AllInst.Text
    $worksheet.Cells.Item($row.Row,8)=$AllReg.Text
    $worksheet.Cells.Item($row.Row,9)=$AllPrice.Text
} 
        }


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
    

# Check if Allocation storage tier is blank
# If it is replace the Allocation values with the Consumption values
$mystring = $AllTier.text
IF([string]::IsNullOrEmpty($mystring)) {            
    Write-Host "Empty Allocation cell found for volume ",$Volume.text," at row ", $row.Row
    $worksheet.Cells.Item($row.Row,5)=$ConsTier.Text
    $worksheet.Cells.Item($row.Row,6)=$ConsSize.Text
    $worksheet.Cells.Item($row.Row,7)=$ConsReg.Text
    $worksheet.Cells.Item($row.Row,8)=$ConsPrice.Text
} 

# Check if Consumption  storage tier is blank
# If it is replace the Consumption values with the Allocation values
$mystring = $ConsTier.text
IF([string]::IsNullOrEmpty($mystring)) {            
    Write-Host "Empty Consumption cell found for volume ",$Volume.text," at row ", $row.Row
    $worksheet.Cells.Item($row.Row,9)=$AllTier.Text
    $worksheet.Cells.Item($row.Row,10)=$AllSize.Text
    $worksheet.Cells.Item($row.Row,11)=$AllReg.Text
    $worksheet.Cells.Item($row.Row,12)=$AllPrice.Text
}
        }

        
    
        


