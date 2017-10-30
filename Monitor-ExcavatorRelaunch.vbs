Option Explicit

Dim oShell, oWMI, oEvents, oReceivedEvent
Set oShell=WScript.CreateObject("WScript.Shell")

Set oWMI=GetObject("winmgmts:\\.\root\CIMV2")
Set oEvents=oWMI.ExecNotificationQuery("SELECT * FROM Win32_ProcessTrace WHERE ProcessName='excavator.exe'")

Do
	Set oReceivedEvent=oEvents.NextEvent
	WScript.Echo oReceivedEvent.ProcessName
Loop