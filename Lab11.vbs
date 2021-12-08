On Error Resume Next
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
If Err.Number <> 0 Then
	WScript.Echo Err.Number & ": " & Err.Description
	WScript.Quit
End If
For Each objPort In objService.ExecQuery("SELECT * FROM Win32_SerialPort")
info= info& "Name of device " & objPort.Caption 'наименование устройства
info= info& "Description of device " & objPort.Description 'описание устройства
info= info& "ID of device " & objPort.DeviceID 'идентификатор устройства
info= info& "PNP ID of device " & objPort.PNPDeviceID 'идентификатор устройства Plug-and-Play
info= info& "Name of computer" & objPort.SystemName 'имя компьютера
Next

For Each objPort In objService.ExecQuery("SELECT * FROM Win32_ParallelPort")
info= info& "Name of device " & objPort.Caption 'наименование устройства
info= info& "Description of device " & objPort.Description 'описание устройства
info= info& "ID of device " & objPort.DeviceID 'идентификатор устройства
info= info& "PNP ID of device " & objPort.PNPDeviceID 'идентификатор устройства Plug-and-Play
info= info& "Name of computer" & objPort.SystemName 'имя компьютера
Next

For Each objObject In objService.ExecQuery("SELECT * FROM Win32_PortResource")
info= info& "Name of device " & objObject.Caption 'наименование устройства
info= info& "Description of device " & objObject.Description 'описание
info= info& "Name of computer " & objObject.CSName 'имя компьютера
info= info& "Start Address " & objObject.StartingAddress 'начальный адрес
info= info& "End Address " & objObject.EndingAddress 'конечный адрес
Next

WScript.Echo info

