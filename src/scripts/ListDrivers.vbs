strComputer = "." 

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_DiskDrive",,48) 
For Each objItem in colItems 
   's = s & "SerialNumber: " & objItem.SerialNumber & vbcrlf 
   's = s & "Caption: " & objItem.Caption
   
   query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + objItem.DeviceID + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition" ' link the physical drives to the partitions
   Set partitions = objWMIService.ExecQuery(query) 
   For Each partition In partitions 
      query = "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + partition.DeviceID + "'} WHERE AssocClass = Win32_LogicalDiskToPartition"  ' link the partitions to the logical disks 
      Set logicalDisks = objWMIService.ExecQuery (query) 
      For Each logicalDisk In logicalDisks      
         WScript.Echo logicalDisk.DeviceID & " --- " & partition.Caption & " --- " & objItem.SerialNumber 
      Next
    Next 
Next


