Option Explicit

'
' THIS SCRIPT IS PROVIDED "AS IS", USE AT YOUR OWN RISK!
' https://github.com/boredazfcuk/mining
'

'----- Initilise Variables -----
Dim Profile, MSIAfterburner, MSIAfterburnerRegPath, sTempFile, oWMI, oFSO, oShell, cProcesses
Dim Process, MSIAfterburnerPath, oFile, RunSilent, UninstallString, MSIAfterburnerInstallPath
Dim sScriptName, ScriptFolder, LogFolder, sLogFile, fLogFile
'----- Initialise CheckMemoryOverclocks Variables -----
Dim nVidiaSMI, QueryMemoryOverclocks, OutputFormat, Count, MemoryOverClocks, aMemoryOverClocks
'----- Initialise Prowl Notification Variables -----
Dim oRegistry, KeyPath, ValueName, ProwlAPIKey, ProwlNotifications, ProwlDisable

Const HKCU=&H80000001
Const OpenAsASCII=0 
Const CreateIfNotExist=1
Const ForAppending=8

'----- Check if launched with profile parameter
If WScript.Arguments.Count > 0 Then
	'----- Confirm parameter is a number -----
	If IsNumeric (WScript.Arguments(0)) Then
		'----- Confirm parameter is between 1 and 5 -----
		If ((WScript.Arguments(0) > 0) And (WScript.Arguments(0) < 6)) Then
			'----- Set profile number to same as parameter -----
			Profile=WScript.Arguments(0)
		'----- If parameter is not between 1 and 5 -----
		Else
			'----- Default to profile 1 -----
			Profile=1
		End If
	'----- If parameter is not a number -----
	Else
		'----- Default to profile 1 -----
		Profile=1
	End If
'----- If parameter is not supplied -----
Else
	'----- Default to profile 1 -----
	Profile=1
End If

'----- Create objects -----
Set oWMI = GetObject("winmgmts:\\localhost\root\CIMV2")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
Set oRegistry=GetObject("winmgmts:\\.\root\default:StdRegProv")

KeyPath="Software\boredazfcuk\mining"
ValueName="ProwlAPIKey"
oRegistry.GetStringValue HKCU, KeyPath, ValueName, ProwlAPIKey
	
If Not IsNull(ProwlAPIKey) Then
	ProwlNotifications=True
End If

'----- Change line below to True to disable Prowl notifications for this script only -----
ProwlDisable=False

'----- Set MSI Afterburner Executable name -----
MSIAfterburner="MSIAfterburner.exe"
MSIAfterburnerRegPath="HKLM\SOFTWARE\WOW6432Node\MSI\Afterburner\InstallPath"
sTempFile = oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName
'----- Get Script Name -----
sScriptName=WScript.ScriptFullName
'----- Get Script Folder -----
Set oFile = oFSO.GetFile(sScriptName)
ScriptFolder=oFSO.GetParentFolderName(oFile)
'----- Set Log Folder Location -----
LogFolder=oFSO.BuildPath(ScriptFolder, "\Logs")
'----- If Log Sub Folder doesn't exist -----
If Not (oFSO.FolderExists(LogFolder)) Then
    '----- Create Log SubFolder-----
    oFSO.CreateFolder(LogFolder)
End If
sLogFile=oFSO.BuildPath(LogFolder, "\Monitor-Overclocks.log")

CheckMemoryOverclocks

Sub CheckMemoryOverclocks
	'----- nVidia SMI query elements -----
	nVidiaSMI="""C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"""
	QueryMemoryOverclocks=" --query-gpu=clocks.max.memory,clocks.current.memory "
	OutputFormat="--format=csv,noheader,nounits"
	
	'----- Query GPU Memory Overclocks -----
	RunSilent = oShell.Run("cmd /c " & nVidiaSMI & QueryMemoryOverclocks & OutputFormat & " > " & sTempFile, 0, True)
	
	Set oFile = oFSO.OpenTextFile(sTempFile, 1)
	'----- Read results for all GPUs into a variable -----
	
	Do While Not oFile.AtEndOfStream
		Count=0
		'----- Read line -----
		MemoryOverClocks = oFile.ReadLine()
		'----- Clean up Line -----
		Trim(MemoryOverClocks)
		'----- Split line into array using comma as separator -----
		aMemoryOverclocks = Split(MemoryOverClocks, ",")
		'----- Check each array element -----
		For Count = 0 To UBound(aMemoryOverclocks)
			'----- Clean up value -----
			aMemoryOverclocks(Count) = Trim(aMemoryOverclocks(Count))
			'----- Add Utilisation to running total -----
		Next
		'----- If GPU's Max Overclock value is greater than the current clock speed -----
		If (aMemoryOverclocks(0) > aMemoryOverclocks(1)) Then
			'----- Reapply the MSI Afterburner profile -----
			ReapplyProfile
		End If
	Loop
	'----- Close Temp file -----
	oFile.Close
	'----- Delete Temp file -----
	oFSO.DeleteFile(sTempFile)
End Sub

Sub ReapplyProfile
	'----- If Prowl Notifications are enabled -----
	If ((ProwlNotifications) And (Not ProwlDisable))Then
		SendProwlNotification "2","Monitor-Overclocks","GPU Overclocks lower than expected - Reapplying MSI Afterburner profile " & Profile
	End If
	'----- Write event to Windows Application Log -----
	oShell.LogEvent 1, "GPU Overclocks lower than expected at " & Now() & " - Reapplying MSI Afterburner profile " & Profile
	Set fLogFile=oFSO.OpenTextFile(sLogFile, ForAppending, CreateIfNotExist, OpenAsASCII)
	'----- Write log to log file -----
	fLogFile.WriteLine ("GPU Overclocks lower than expected at " & Now() & " - Reapplying MSI Afterburner profile " & Profile)
	'----- Close log file -----
	fLogFile.Close
	'----- Grab MSI Afterburner Process details -----
	Set cProcesses = oWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Caption='" & MSIAfterburner & "'")
	'----- If MSI Afterburner is running -----
	If cProcesses.Count > 0 Then
		'----- Check MSI Afterburner process details -----
		For Each Process In cProcesses
			'----- Grab command line used to launch MSIAfterburner -----
			MSIAfterburnerPath = Process.CommandLine
		Next
		'----- Apply MSI Afterburner profile -----
		RunSilent=oShell.Run(MSIAfterburnerPath & " -profile" & Profile, 0, False)
	Else
		'----- Check MSI Afterburner Install Location -----
	    MSIAfterburnerInstallPath=oShell.RegRead(MSIAfterburnerRegPath)
	    '----- Run MSI Afterburner applying specified profile -----
	   	RunSilent=oShell.Run(Chr(34) & MSIAfterburnerInstallPath & Chr(34) & " /s -profile" & Profile, 0, False)
	   	'----- Close Temp file -----
		oFile.Close
		'----- Delete Temp file -----
		oFSO.DeleteFile(sTempFile)
	   	'----- Quit out, rather than check the rest of the overclocks, as they should be good
	   	WScript.Quit(0)
	End If
End Sub

Sub SendProwlNotification(Priority, Application, Description)
	Dim oHTTP
	Set oHTTP=CreateObject("Microsoft.XMLHTTP")  
	oHTTP.Open "Get", "https://prowl.weks.net/publicapi/add?" & "apikey=" & ProwlAPIKey & "&priority=" & Priority & "&application=" & Application & "&event=" & Date() & " " & Time()  & "&description=" & Description ,false  
	oHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"  
	oHTTP.Send  
End Sub