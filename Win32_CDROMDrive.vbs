Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
For Each objDisk In objService.ExecQuery("SELECT * FROM Win32_CDROMDrive")
WScript.Echo objDisk.SystemName 'имя компьютера
WScript.Echo objDisk.Caption 'наименование устройства
WScript.Echo objDisk.Description 'описание устройства
WScript.Echo objDisk.DeviceID 'идентификатор устройства
WScript.Echo objDisk.Manufacturer 'производитель
WScript.Echo objDisk.Id 'drive letter
WScript.Echo objDisk.Size 'размер диска
WScript.Echo objDisk.VolumeName 'метка тома
WScript.Echo objDisk.VolumeSerialNumber 'серийный номер тома
Next