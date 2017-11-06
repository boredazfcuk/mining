Option Explicit

'
' THIS SCRIPT IS PROVIDED "AS IS", USE AT YOUR OWN RISK!
' https://github.com/boredazfcuk/mining
'

'----- Initialise Variables -----
Dim oShell, oFSO, nVidiaSMI, QueryPowerLimit, OutputFormat, sTempFile, CurrentFolder, LogFolder
Dim sLogFile, RunSilent, oFile, GPUPowerLimit, aGPUPowerLimit, fLogFile, Count, Total
Dim oRegistry, KeyPath, ValueName, ProwlAPIKey, ProwlNotifications, ProwlDisable

'----- Create objects -----
Set oShell=CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")
Set oRegistry=GetObject("winmgmts:\\.\root\default:StdRegProv")

'----- Create Constants -----
Const ForAppending = 8
Const CreateIfNotExist = 1
Const OpenAsASCII = 0
Const MaxPower = 880
Const HKCU=&H80000001

'----- Set variables -----
nVidiaSMI="""C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"""
QueryPowerLimit=" --query-gpu=power.limit "
OutputFormat="--format=csv,noheader,nounits"
sTempFile=oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName
KeyPath="Software\boredazfcuk\mining"
ValueName="ProwlAPIKey"
oRegistry.GetStringValue HKCU, KeyPath, ValueName, ProwlAPIKey, ProwlDisable
	
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
'----- Set full log file path -----
sLogFile =  oFSO.BuildPath(LogFolder, "\Monitor-Power.log")

'----- Query GPU Power Limit -----
RunSilent = oShell.Run("cmd /c " & nVidiaSMI & QueryPowerLimit & OutputFormat & " > " & sTempFile, 0, True)
'----- Target Temp File -----
Set oFile = oFSO.OpenTextFile(sTempFile, 1)
'----- Read whole file to variable -----
GPUPowerLimit = oFile.ReadAll
'----- Clean up query results variable -----
Trim(GPUPowerLimit)
'----- Close Temp file -----
oFile.Close
'----- Delete Temp file -----
oFSO.DeleteFile(sTempFile)

'----- Split each line into separate array element -----
aGPUPowerLimit=Split(GPUPowerLimit,vbCrLf)

'----- Loop through array elements -----
For Count = 0 To UBound(aGPUPowerLimit)-1
	'----- Clean up element's value -----
	aGPUPowerLimit(Count) = Trim(aGPUPowerLimit(Count))
	'----- Add element's value to running total -----
	Total = Total + Int(aGPUPowerLimit(Count))
Next

'----- If running total is more than maximum power limit -----
If Total > MaxPower Then
	'----- If Prowl Notifications are enabled -----
	If ((ProwlNotifications) And (Not ProwlDisable))Then
		SendProwlNotification "2","Monitor-Power","GPUs drawing too much power - Shutting down."
	End If
	'----- Target log file -----
	Set fLogFile = oFSO.OpenTextFile(sLogFile, ForAppending, CreateIfNotExist, OpenAsASCII)
	'----- Log event to Windows Application log -----
	oShell.LogEvent 1, "GPUs drawing too much power " & Now() & " - Shutting down."
	'----- Log to Log file
	fLogFile.WriteLine ("GPUs drawing too much power " & Now() & " - Shutting down.")
	'----- Close Log File
	fLogFile.Close
	'----- Reboot computer -----
	RunSilent=oShell.Run("shutdown /f /s /t 60", , True)
End If

Sub SendProwlNotification(Priority, Application, Description)
	Dim oHTTP
	Set oHTTP=CreateObject("Microsoft.XMLHTTP")  
	oHTTP.Open "Get", "https://prowl.weks.net/publicapi/add?" & "apikey=" & ProwlAPIKey & "&priority=" & Priority & "&application=" & Application & "&event=" & Date() & " " & Time()  & "&description=" & Description ,false  
	oHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"  
	oHTTP.Send  
End Sub