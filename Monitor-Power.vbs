Option Explicit

Dim nVidiaSMI, OutputFormat, oShell, oFSO, sTempFile, Return, oFile, oGPUPowerLimit, GPUPowerLimit, aGPUPowerLimit, sLogFile, fLogFile, i, Total

nVidiaSMI="""C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"""
OutputFormat="--format=csv,noheader,nounits"

Set oShell=CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")
Const ForAppending = 8
Const CreateIfNotExist = 1
Const OpenAsASCII = 0
sTempFile=oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName
sLogFile = "C:\Scripts\Logs\Monitor-Power.log"

Return = oShell.Run("cmd /c " & nVidiaSMI & " --query-gpu=power.limit " & OutputFormat & " > " & sTempFile, 0, True)
Set oFile = oFSO.OpenTextFile(sTempFile, 1)
GPUPowerLimit = oFile.ReadAll
Trim(GPUPowerLimit)
oFile.Close
oFSO.DeleteFile(sTempFile)

aGPUPowerLimit=Split(GPUPowerLimit,vbCrLf)
For i = 0 To UBound(aGPUPowerLimit)-1
	aGPUPowerLimit(i) = Trim(aGPUPowerLimit(i))
	Total = Total + Int(aGPUPowerLimit(i))
Next

If Total > 880 Then
	Set fLogFile = oFSO.OpenTextFile(sLogFile, ForAppending, CreateIfNotExist, OpenAsASCII)
	fLogFile.WriteLine ("GPUs drawing too much power " & Now() & " - Shutting down.")
	fLogFile.Close
	Return=oShell.Run("%comspec% /c shutdown /f /s /t 60", , True)
End If