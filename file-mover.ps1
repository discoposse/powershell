# Script to move assessment files from a starting folder to a file structure for each group
# Use Ryan's tbutil cloud_migration_plans.js to download the 3 CSV files to a starting folder
# This script will create subfolders based on the leading group name and copy the 3 files
# into each new folder
# Once in a nice structure, can use auto-assess-folder.ps1 to create the master spreadsheet
# Richard Stinton - April 2019

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

# Get the folder to work on
$folder = Get-Folder

$path = $folder + "\*breakdown*.csv"

$files = get-childitem -Path $path

foreach ($file in $files) {
    $groupname,$rest = $file.BaseName.Split('_')
    $newfoldername = $folder + "\" + $groupname
    $newfoldername
    New-Item -Path $newfoldername -ItemType Directory
    $stuff = $folder + "\" + $groupname + "_*.csv"
    move-item $stuff $newfoldername
    
    }