Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
If Err.Number <> 0 Then
WScript.Echo Err.Number & ": " & Err.Description
WScript.Quit
End If
For Each objDisk In objService.ExecQuery("SELECT * FROM Win32_DiskDrive")
WScript.Echo objDisk.SystemName 'имя компьютера
WScript.Echo objDisk.Caption 'наименование устройства
WScript.Echo objDisk.Model 'модель, указанная производителем
WScript.Echo objDisk.Description 'описание устройства
WScript.Echo objDisk.DeviceID 'идентификатор устройства
WScript.Echo objDisk.PNPDeviceID 'идентификатор устройства Plug-and-Play
WScript.Echo objDisk.Manufacturer 'производитель
WScript.Echo objDisk.Index 'номер диска (если 0xFF - не отображает физический диск)
WScript.Echo objDisk.InterfaceType 'тип интерфейса (IDE, SCSI)
WScript.Echo objDisk.MediaType 'тип носителя (Removable media, Fixed hard disk
WScript.Echo objDisk.SCSIBus 'номер шины SCSI
WScript.Echo objDisk.SCSILogicalUnit 'номер SCSI устройства
WScript.Echo objDisk.SCSIPort 'номер порта SCSI
WScript.Echo objDisk.SCSITargetId 'идентификационный номер SCSI
WScript.Echo objDisk.TotalHeads 'количество головок
WScript.Echo objDisk.BytesPerSector 'количество байт в секторе
WScript.Echo objDisk.SectorsPerTrack 'количество секторов на дорожке
WScript.Echo objDisk.TracksPerCylinder 'количество дорожек в цилиндре
WScript.Echo objDisk.TotalCylinders 'количество цилиндров
WScript.Echo objDisk.TotalSectors 'общее количество секторов
WScript.Echo objDisk.TotalTracks 'общее количество дорожек
WScript.Echo objDisk.Size 'размер диска (по количеству цилиндров, дорожек, секторов иразмеру сектора)
WScript.Echo objDisk.Partitions 'количество разделов на диске
WScript.Echo
Next
54585272525245245254