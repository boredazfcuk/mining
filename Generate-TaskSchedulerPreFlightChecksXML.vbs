Option Explicit

'
' THIS SCRIPT IS PROVIDED "AS IS", USE AT YOUR OWN RISK!
' https://github.com/boredazfcuk/mining
'

'----- Initialise Variables -----
Dim oFSO, oFile, sScriptName, ScriptFolder, sXMLFile, fXMLFile, VBSFile, Count, oShell, RunSilent

'----- Create Objects -----
Set oFSO=CreateObject("Scripting.FileSystemObject")
Set oShell=CreateObject("WScript.Shell")

'----- Set Constants -----
Const OpenAsASCII=0 
Const CreateIfNotExist=1
Const ForAppending=8

'----- Get Script Name -----
sScriptName=WScript.ScriptFullName
'----- Get Script Folder -----
Set oFile = oFSO.GetFile(sScriptName)
ScriptFolder=oFSO.GetParentFolderName(oFile)
'----- Set path to Tash Scheduler XML file to create -----
sXMLFile=oFSO.BuildPath(ScriptFolder, "\Monitor-PreFlightChecks.xml")

'----- If XML File has already been created -----
If oFSO.FileExists(sXMLFile) Then
	'----- Delete it -----
	oFSO.DeleteFile sXMLFile
End If

'----- Set XML file Target -----
Set fXMLFile = oFSO.OpenTextFile(sXMLFile, ForAppending, CreateIfNotExist, OpenAsASCII)

'----- Create array with the name of each script to add to task scheduler in it -----
VBSFile="Monitor-PreFlightChecks.vbs"

'----- Kick out the XML header -----
fXMLFile.WriteLine ("<?xml version=""1.0"" encoding=""UTF-16""?>")
fXMLFile.WriteLine ("<Task version=""1.2"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">")
fXMLFile.WriteLine ("  <RegistrationInfo>")
fXMLFile.WriteLine ("    <URI>\Monitor-PreFlightChecks</URI>")
fXMLFile.WriteLine ("  </RegistrationInfo>")
fXMLFile.WriteLine ("  <Triggers>")
fXMLFile.WriteLine ("    <LogonTrigger>")
fXMLFile.WriteLine ("      <Enabled>true</Enabled>")
fXMLFile.WriteLine ("    </LogonTrigger>")
fXMLFile.WriteLine ("  </Triggers>")
fXMLFile.WriteLine ("  <Principals>")
fXMLFile.WriteLine ("    <Principal id=""Author"">")
fXMLFile.WriteLine ("      <LogonType>InteractiveToken</LogonType>")
fXMLFile.WriteLine ("      <RunLevel>HighestAvailable</RunLevel>")
fXMLFile.WriteLine ("    </Principal>")
fXMLFile.WriteLine ("  </Principals>")
fXMLFile.WriteLine ("  <Settings>")
fXMLFile.WriteLine ("    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>")
fXMLFile.WriteLine ("    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>")
fXMLFile.WriteLine ("    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>")
fXMLFile.WriteLine ("    <AllowHardTerminate>true</AllowHardTerminate>")
fXMLFile.WriteLine ("    <StartWhenAvailable>true</StartWhenAvailable>")
fXMLFile.WriteLine ("    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>")
fXMLFile.WriteLine ("    <IdleSettings>")
fXMLFile.WriteLine ("      <StopOnIdleEnd>true</StopOnIdleEnd>")
fXMLFile.WriteLine ("      <RestartOnIdle>false</RestartOnIdle>")
fXMLFile.WriteLine ("    </IdleSettings>")
fXMLFile.WriteLine ("    <AllowStartOnDemand>true</AllowStartOnDemand>")
fXMLFile.WriteLine ("    <Enabled>true</Enabled>")
fXMLFile.WriteLine ("    <Hidden>false</Hidden>")
fXMLFile.WriteLine ("    <RunOnlyIfIdle>false</RunOnlyIfIdle>")
fXMLFile.WriteLine ("    <WakeToRun>false</WakeToRun>")
fXMLFile.WriteLine ("    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>")
fXMLFile.WriteLine ("    <Priority>7</Priority>")
fXMLFile.WriteLine ("  </Settings>")
fXMLFile.WriteLine ("  <Actions Context=""Author"">")
fXMLFile.WriteLine ("    <Exec>")
fXMLFile.WriteLine ("      <Command>wscript.exe</Command>")
fXMLFile.WriteLine ("      <Arguments>//nologo " & Chr(34) & ScriptFolder & "\" & VBSFile & Chr(34) & "</Arguments>")
fXMLFile.WriteLine ("    </Exec>")
fXMLFile.WriteLine ("  </Actions>")
fXMLFile.WriteLine ("</Task>")

'----- Close XML File so a lock condition doesn't exist -----
fXMLFile.Close

'----- Import Task to Task Scheduler -----
oShell.Run "schtasks.exe /Create /XML " & Chr(34) & sXMLFile & Chr(34) & " /TN Monitor-Critical", 0, True