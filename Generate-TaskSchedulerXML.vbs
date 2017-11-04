Option Explicit

'
' THIS SCRIPT IS PROVIDED "AS IS", USE AT YOUR OWN RISK!
' https://github.com/boredazfcuk/mining
'

'----- Initialise Variables -----
Dim oFSO, CurrentFolder, sXMLFile, fXMLFile, aVBSFiles(6), VBSFile, Count, oShell, RunSilent

'----- Create Objects -----
Set oFSO=CreateObject("Scripting.FileSystemObject")
Set oShell=CreateObject("WScript.Shell")

'----- Set Constants -----
Const OpenAsASCII=0 
Const CreateIfNotExist=1
Const ForAppending=8

'----- Get Script Folder -----
CurrentFolder=oFSO.GetAbsolutePathName(".")
'----- Set path to Tash Scheduler XML file to create -----
sXMLFile=oFSO.BuildPath(CurrentFolder, "\Monitor-Critical.xml")

'----- If XML File has already been created -----
If oFSO.FileExists(sXMLFile) Then
	'----- Delete it -----
	oFSO.DeleteFile sXMLFile
End If

'----- Set XML file Target -----
Set fXMLFile = oFSO.OpenTextFile(sXMLFile, ForAppending, CreateIfNotExist, OpenAsASCII)

'----- Create array with the name of each script to add to task scheduler in it -----
aVBSFiles(0)="Monitor-GPUTotal.vbs"
aVBSFiles(1)="Monitor-NetworkConnection.vbs"
aVBSFiles(2)="Monitor-NiceHash.vbs"
aVBSFiles(3)="Monitor-OverClocks.vbs"
aVBSFiles(4)="Monitor-Power.vbs"
aVBSFiles(5)="Monitor-PowerLevels.vbs"
aVBSFiles(6)="Monitor-PRTGProbeService.vbs"

'----- Kick out the XML header -----
fXMLFile.WriteLine ("<?xml version=""1.0"" encoding=""UTF-16""?>")
fXMLFile.WriteLine ("<Task version=""1.2"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">")
fXMLFile.WriteLine ("  <RegistrationInfo>")
fXMLFile.WriteLine ("    <Date>2017-09-13T00:40:54.4035683</Date>")
fXMLFile.WriteLine ("    <URI>\Monitor-Critical</URI>")
fXMLFile.WriteLine ("  </RegistrationInfo>")
fXMLFile.WriteLine ("  <Triggers>")
fXMLFile.WriteLine ("    <BootTrigger>")
fXMLFile.WriteLine ("      <Repetition>")
fXMLFile.WriteLine ("        <Interval>PT1M</Interval>")
fXMLFile.WriteLine ("        <StopAtDurationEnd>false</StopAtDurationEnd>")
fXMLFile.WriteLine ("      </Repetition>")
fXMLFile.WriteLine ("      <Enabled>true</Enabled>")
fXMLFile.WriteLine ("      <Delay>PT5M</Delay>")
fXMLFile.WriteLine ("    </BootTrigger>")
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
'----- Loop through the array to create the list of programs to start -----
For Each VBSFile In aVBSFiles
	fXMLFile.WriteLine ("    <Exec>")
	fXMLFile.WriteLine ("      <Command>wscript.exe</Command>")
	fXMLFile.WriteLine ("      <Arguments>//nologo " & Chr(34) & CurrentFolder & "\" & aVBSFiles(Count) & Chr(34) & "</Arguments>")
	fXMLFile.WriteLine ("    </Exec>")
	Count=Count+1
Next
'----- Write XML footer -----
fXMLFile.WriteLine ("  </Actions>")
fXMLFile.WriteLine ("</Task>")

'----- Close XML File so a lock condition doesn't exist -----
fXMLFile.Close

'----- Import Task to Task Scheduler -----
oShell.Run "schtasks.exe /Create /XML " & Chr(34) & sXMLFile & Chr(34) & " /TN Monitor-Critical", 0, True