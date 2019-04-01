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

#$folder = Get-Folder
$folder = 'Z:\data\customers\Microsoft\pstest\Large Assessment'
$path = $folder + "\*breakdown*.csv"

$files = get-childitem -Path $path

foreach ($file in $files) {
    $file.BaseName
    $groupname,$rest = $file.BaseName.Split('_')
    $newfoldername = $folder + "\" + $groupname
    $newfoldername
    New-Item -Path $newfoldername -ItemType Directory
    $stuff = $folder + "\" + $groupname + "_*.csv"
    $stuff
    move-item $stuff $newfoldername
    
    }