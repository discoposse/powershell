##
$id=0
$data = Import-CSV "Z:\data\powershell\cmdb_civica.csv"
$outfile = "Z:\data\powershell\cmdb_civica.txt"
$stuff = '{"hosts": ['
$stuff | out-file $outfile
foreach ($row in $data) {
    $id++
    $idname = "id"+$id
    write-host $id
    $hostname = $row.displayName
    $IPaddress = $row.ipAddress
    $OS = $row.osName
    $cores = $row.numOfCPU
    $cpu_speed = $row.cpuSpeedMHz
    $mem = $row.memsizeGB
    $disk = $row.diskSizeGB
    write-host "Working on host ",$hostname," with IP ",$IPaddress
    write-host "CPU speed ",$cpu_speed," Cores ",$cores," RAM ",$mem," Disk ",$disk
    $stuff = '{"entityId" : "'+$idname+'","displayName": "'+$hostname+'","osName" : "'+$os+'", "numOfCPU" : "'+$cores+'", "cpuSpeedMhz" : "'+$cpu_speed+ '", "memsizeGB" : "'+$mem+'", "diskSizeGB" : "'+$disk+'", "ipAddresses": ["'+$IPaddress+'"]},'
    $stuff | out-file -append $outfile
}
$stuff=']}'
$stuff | out-file -append $outfile
$stuff ='Please remove comma from end of last but one line....and then remove this line'
$stuff | out-file -append $outfile