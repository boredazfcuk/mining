Option Explicit

'
' THIS SCRIPT IS PROVIDED "AS IS", USE AT YOUR OWN RISK!
' https://github.com/boredazfcuk/mining
'

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