On Error Resume Next
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
If Err.Number <> 0 Then
	WScript.Echo Err.Number & ": " & Err.Description
	WScript.Quit
End If
For Each objPnP In objService.ExecQuery("SELECT * FROM Win32_PnPEntity")
	WScript.Echo objPnP.Name 'наименование устройства
	WScript.Echo objPnP.Description 'описание устройства
	WScript.Echo objPnP.Manufacturer 'производитель
	WScript.Echo objPnP.PNPDeviceID 'идентификатор логического устройства
	WScript.Echo
Next