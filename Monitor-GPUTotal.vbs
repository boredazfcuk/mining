Option Explicit

Dim nVidiaSMI, QueryTotalMem, OutputFormat, oShell, oFSO, sTempFile, Return, oFile, oGPUMemoryTotal, GPUMemoryTotal, aGPUMemoryTotal, sLogFile, fLogFile, i, Total, GPUNumber

nVidiaSMI="""C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"""
QueryTotalMem=" --query-gpu=memory.total "
OutputFormat="--format=csv,noheader,nounits"

Set oShell=CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")
Const ForAppending = 8
Const CreateIfNotExist = 1
Const OpenAsASCII = 0
sTempFile=oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName
sLogFile = "C:\Scripts\Logs\Monitor-GPUsTotal.log"
GPUNumber=0

Return = oShell.Run("cmd /c " & nVidiaSMI & QueryTotalMem & OutputFormat & " > " & sTempFile, 0, True)
Set oFile = oFSO.OpenTextFile(sTempFile, 1)
GPUMemoryTotal = oFile.ReadAll
Trim(GPUMemoryTotal)
oFile.Close
oFSO.DeleteFile(sTempFile)

aGPUMemoryTotal=Split(GPUMemoryTotal,vbCrLf)
For i = 0 To UBound(aGPUMemoryTotal)-1
	aGPUMemoryTotal(i) = Trim(aGPUMemoryTotal(i))
	If aGPUMemoryTotal(i) = "[Unknown Error]" Then
		RebootComputer
	End If
	GPUNumber=GPUNumber+1
Next

Sub RebootComputer
	Set fLogFile = oFSO.OpenTextFile(sLogFile, ForAppending, CreateIfNotExist, OpenAsASCII)
	fLogFile.WriteLine ("GPU#" & GPUNumber & " Failed at " & Now() & " - Rebooting.")
	fLogFile.Close
	Return=oShell.Run("%comspec% /c shutdown /f /r /t 60", , True)
End Sub