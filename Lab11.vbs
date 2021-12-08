On Error Resume Next
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
If Err.Number <> 0 Then
	WScript.Echo Err.Number & ": " & Err.Description
	WScript.Quit
End If
For Each objPort In objService.ExecQuery("SELECT * FROM Win32_ParallelPort")
info= info& "Name of device " & objPort.Caption 'наименование устройства
info= info& "Description of device " & objPort.Description 'описание устройства
info= info& "ID of device " & objPort.DeviceID 'идентификатор устройства
info= info& "PNP ID of device " & objPort.PNPDeviceID 'идентификатор устройства Plug-and-Play
info= info& "Name of computer" & objPort.SystemName 'имя компьютера
Next
WScript.Echo info
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_PortResource")
	WScript.Echo objObject.Caption 'наименование устройства
	WScript.Echo objObject.Description 'описание
	WScript.Echo objObject.CSName 'имя компьютера
	WScript.Echo objObject.StartingAddress 'начальный адрес
	WScript.Echo objObject.EndingAddress 'конечный адрес
	WScript.Echo
Next
For Each objPort In objService.ExecQuery("SELECT * FROM Win32_ParallelPort")
	WScript.Echo objPort.Caption 'наименование устройства
	WScript.Echo objPort.Description 'описание устройства
	WScript.Echo objPort.DeviceID 'идентификатор устройства
	WScript.Echo objPort.PNPDeviceID 'идентификатор устройства Plug-and-Play
	WScript.Echo objPort.SystemName 'имя компьютера
Next