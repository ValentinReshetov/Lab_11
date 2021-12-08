On Error Resume Next
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
If Err.Number <> 0 Then
	WScript.Echo Err.Number & ": " & Err.Description
	WScript.Quit
End If
For Each objPort In objService.ExecQuery("SELECT * FROM Win32_SerialPort, Win32_PortResource")
	WScript.Echo objPort.Caption 'наименование устройства
	WScript.Echo objPort.Description 'описание устройства
	WScript.Echo objPort.DeviceID 'идентификатор устройства
	WScript.Echo objPort.PNPDeviceID 'идентификатор устройства Plug-and-Play
	WScript.Echo objPort.SystemName 'имя компьютера
Next
For Each objPort In objService.ExecQuery("SELECT * FROM Win32_PortResource")
	WScript.Echo objObject.Caption 'наименование устройства
	WScript.Echo objObject.Description 'описание
	WScript.Echo objObject.CSName 'имя компьютера
	WScript.Echo objObject.StartingAddress 'начальный адрес
	WScript.Echo objObject.EndingAddress 'конечный адрес
Next