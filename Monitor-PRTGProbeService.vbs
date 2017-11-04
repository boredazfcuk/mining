Option Explicit

'----- Initialise Variables -----
Dim oService

'----- Set Object relating to PRTG service -----
Set oService = GetObject("WinNT://./PRTGProbeService,service")
'WScript.Echo oService.Status

'----- If service is stopped -----
If oService.Status=1 Then
'----- Start Service -----
	oService.Start
End If