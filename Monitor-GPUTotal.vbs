Option Explicit

'
' THIS SCRIPT IS PROVIDED "AS IS", USE AT YOUR OWN RISK!
' https://github.com/boredazfcuk/mining
'

'----- Define variables -----
Dim nVidiaSMI, QueryTotalMem, OutputFormat, RegKey, oShell, oFSO, oRegistry, sTempFile, KeyPath
Dim ValueName, ProwlAPIKey, ProwlNotifications, ProwlDisable, CurrentFolder, LogFolder, RunSilent
Dim oFile, oGPUMemoryTotal, GPUMemoryTotal, aGPUMemoryTotal, sLogFile, fLogFile, Count, Total

nVidiaSMI="""C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"""
QueryTotalMem=" --query-gpu=memory.total "
OutputFormat="--format=csv,noheader,nounits"

Set oShell=CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")
Set oRegistry=GetObject("winmgmts:\\.\root\default:StdRegProv")

Const ForAppending = 8
Const CreateIfNotExist = 1
Const OpenAsASCII = 0
Const HKCU=&H80000001
sTempFile=oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName
KeyPath="Software\boredazfcuk\mining"
ValueName="ProwlAPIKey"
oRegistry.GetStringValue HKCU, KeyPath, ValueName, ProwlAPIKey
	
If Not IsNull(ProwlAPIKey) Then
	ProwlNotifications=True
End If

'----- Change line below to True to disable Prowl notifications for this script only -----
ProwlDisable=False

'----- Get Script Folder -----
CurrentFolder = oFSO.GetAbsolutePathName(".")
LogFolder = oFSO.BuildPath(CurrentFolder, "\Logs")
'----- If Log Sub Folder doesn't exist -----
If Not (oFSO.FolderExists(LogFolder)) Then
    '----- Create Log SubFolder-----
    oFSO.CreateFolder(LogFolder)
End If
sLogFile =  oFSO.BuildPath(LogFolder, "\Monitor-GPUsTotal.log")

'----- nVidia SMI query elements -----
RunSilent = oShell.Run("cmd /c " & nVidiaSMI & QueryTotalMem & OutputFormat & " > " & sTempFile, 0, True)
Set oFile = oFSO.OpenTextFile(sTempFile, 1)
'----- Read results to variable -----
GPUMemoryTotal = oFile.ReadAll
'----- Clean up variable -----
Trim(GPUMemoryTotal)
'----- Close Temp file -----
oFile.Close
'----- Delete Temp file -----
oFSO.DeleteFile(sTempFile)

'----- Split query results variable by line to make array -----
aGPUMemoryTotal=Split(GPUMemoryTotal,vbCrLf)
'----- For each array element -----
For Count = 0 To UBound(aGPUMemoryTotal)-1
	'----- Clean up each array element -----
	aGPUMemoryTotal(Count) = Trim(aGPUMemoryTotal(Count))
	'----- If the element's value contains [ and ] then -----
	If ((InStr(aGPUMemoryTotal(Count),"[") And InStr(aGPUMemoryTotal(Count),"]")) Or InStr(aGPUMemoryTotal(Count),"Reboot")) Then
	'----- Go to the reboot computer sub routine -----
		RebootComputer
	End If
Next

Sub RebootComputer
	If ((ProwlNotifications) And (Not ProwlDisable))Then
		SendProwlNotification "2","Monitor-GPUTotal","GPU#" & Count & " Failed - Rebooting."
	End If
	'----- Write event to Windows Application Log -----
	oShell.LogEvent 1, "GPU#" & Count & " Failed at " & Now() & " - Rebooting."
	Set fLogFile = oFSO.OpenTextFile(sLogFile, ForAppending, CreateIfNotExist, OpenAsASCII)
	'----- Write log to log file -----
	fLogFile.WriteLine ("GPU#" & Count & " Failed at " & Now() & " - Rebooting.")
	'----- Close log file -----
	fLogFile.Close
	'----- Reboot Computer -----
	RunSilent=oShell.Run("%comspec% /c shutdown /f /r /t 60", , True)
End Sub

Sub SendProwlNotification(Priority, Application, Description)
	Dim oHTTP
	Set oHTTP=CreateObject("Microsoft.XMLHTTP")  
	oHTTP.Open "Get", "https://prowl.weks.net/publicapi/add?" & "apikey=" & ProwlAPIKey & "&priority=" & Priority & "&application=" & Application & "&event=" & Date() & " " & Time()  & "&description=" & Description ,false  
	oHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"  
	oHTTP.Send  
End Sub