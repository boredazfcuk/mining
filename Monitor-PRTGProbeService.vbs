Option Explicit

Dim oService

Set oService = GetObject("WinNT://./PRTGProbeService,service")
'WScript.Echo oService.Status
 
If oService.Status=1 Then
oService.Start
End If