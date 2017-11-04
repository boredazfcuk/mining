Option Explicit

'
' THIS SCRIPT IS PROVIDED "AS IS", USE AT YOUR OWN RISK!
' https://github.com/boredazfcuk/mining
'

'----- Create Variables -----
Dim oShell, oWMI, oEvents, oReceivedEvent

'----- Create Shell Object ------
Set oShell=WScript.CreateObject("WScript.Shell")

'----- Check if running under "WScript"
If InStr(1, WScript.FullName, "WScript.exe", vbTextCompare) <> 0 Then
	'----- If it is, relaunch uncer cscript.exe -----
        oShell.Run "%comspec% /c cscript /nologo """ & WScript.ScriptFullName & """", 1, False
End If

'----- Create WMI Object -----
Set oWMI=GetObject("winmgmts:\\.\root\CIMV2")

'----- Create WMI Trace Object Events List for excavator.exe -----
Set oEvents=oWMI.ExecNotificationQuery("SELECT * FROM Win32_ProcessTrace WHERE ProcessName='excavator.exe'")

'----- Keep checking the process trace -----
Do
	'----- If new event appears -----
	Set oReceivedEvent=oEvents.NextEvent
	'----- Echo the process name -----
	WScript.Echo oReceivedEvent.ProcessName
Loop