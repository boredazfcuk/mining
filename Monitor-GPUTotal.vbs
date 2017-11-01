Option Explicit

'----- Define variables -----
Dim nVidiaSMI, QueryTotalMem, OutputFormat, oShell, oFSO, sTempFile, CurrentFolder, LogFolder, RunSilent, oFile, oGPUMemoryTotal, GPUMemoryTotal, aGPUMemoryTotal, sLogFile, fLogFile, Count, Total

nVidiaSMI="""C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"""
QueryTotalMem=" --query-gpu=memory.total "
OutputFormat="--format=csv,noheader,nounits"

Set oShell=CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")
Const ForAppending = 8
Const CreateIfNotExist = 1
Const OpenAsASCII = 0
sTempFile=oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName

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
	If (InStr("[",aGPUMemoryTotal(Count)) And InStr("]",aGPUMemoryTotal(Count))) Then
	'----- Go to the reboot computer sub routine -----
		RebootComputer
	End If
Next

Sub RebootComputer
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
