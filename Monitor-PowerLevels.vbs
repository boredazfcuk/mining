Option Explicit

'
' THIS SCRIPT IS PROVIDED "AS IS", USE AT YOUR OWN RISK!
' https://github.com/boredazfcuk/mining
'

'----- Initialise Variables -----
Dim oShell, oFSO, PowerLimit(5), nVidiaSMI, QueryPowerLimit, OutputFormat, sTempFile, CurrentFolder, LogFolder, Return, oFile, oGPUPowerLevels, GPUPowerLevels, aGPUPowerLevels, sLogFile, fLogFile, Count, Total
'----- Initialise Prowl Notification Variables -----
Dim oRegistry, KeyPath, ValueName, ProwlAPIKey, ProwlNotifications, ProwlDisable

'----- Set Constants -----
Const ForAppending = 8
Const CreateIfNotExist = 1
Const OpenAsASCII = 0
Const HKCU=&H80000001

'----- Create Objects -----
Set oShell=CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")
Set oRegistry=GetObject("winmgmts:\\.\root\default:StdRegProv")

'----- Set Variables -----
PowerLimit(0) = "124.00"
PowerLimit(1) = "120.00"
PowerLimit(2) = "120.00"
PowerLimit(3) = "124.00"
PowerLimit(4) = "120.00"
PowerLimit(5) = "120.00"
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
sLogFile =  oFSO.BuildPath(LogFolder, "\Monitor-PowerLevels.log")

'----- Query GPUs power levels -----
Return = oShell.Run("cmd /c " & nVidiaSMI & QueryPowerLimit & OutputFormat & " > " & sTempFile, 0, True)

'----- Target Temp files -----
Set oFile = oFSO.OpenTextFile(sTempFile, 1)
'----- Read whole file into variable -----
GPUPowerLevels = oFile.ReadAll
'----- Clean up query results variable -----
Trim(GPUPowerLevels)
'----- Close Temp File -----
oFile.Close
'----- Delete Temp File -----
oFSO.DeleteFile(sTempFile)

'----- Split each line into separate array element -----
aGPUPowerLevels=Split(GPUPowerLevels,vbCrLf)
'----- Loop through array elements -----
For Count = 0 To UBound(aGPUPowerLevels)-1
	'----- Clean up element's value -----
	aGPUPowerLevels(Count) = Trim(aGPUPowerLevels(Count))
	'----- Check power level -----
	If aGPUPowerLevels(Count) = PowerLimit(Count) Then
		'----- If OK, display message -----
		'WScript.Echo "GPU#" & i & " Power limit " & aGPUPowerLevels(i) & " good"
	Else
		'----- If Prowl Notifications are enabled -----
		If ((ProwlNotifications) And (Not ProwlDisable))Then
			SendProwlNotification "2","Monitor-PowerLevels","GPU#" & Count & " Power limit " & aGPUPowerLevels(Count) & " bad, changing to " & PowerLimit(Count)
		End If
		'----- If bad, write error to Windows Event Viewer Application Log -----
		oShell.LogEvent 1, "GPU#" & Count & " Power limit " & aGPUPowerLevels(Count) & " bad, changing to " & PowerLimit(Count) & " @ " & Now()
		'----- Target Log File -----
		Set fLogFile = oFSO.OpenTextFile(sLogFile, ForAppending, CreateIfNotExist, OpenAsASCII)
		'----- Write error message to log file -----
		fLogFile.WriteLine ("GPU#" & Count & " Power limit " & aGPUPowerLevels(Count) & " bad, changing to " & PowerLimit(Count) & " @ " & Now())
		'----- Close log file -----
		fLogFile.Close
		'----- Set Power Limit to value stored in PowerLimit Variable -----
		Return = oShell.Run("cmd /c " & nVidiaSMI & " -i " & Count & " -pl " & PowerLimit(Count), 0, True)		
	End If
Next

Sub SendProwlNotification(Priority, Application, Description)
	Dim oHTTP
	Set oHTTP=CreateObject("Microsoft.XMLHTTP")  
	oHTTP.Open "Get", "https://prowl.weks.net/publicapi/add?" & "apikey=" & ProwlAPIKey & "&priority=" & Priority & "&application=" & Application & "&event=" & Date() & " " & Time()  & "&description=" & Description ,false  
	oHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"  
	oHTTP.Send  
End Sub