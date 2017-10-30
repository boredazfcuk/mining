Option Explicit

Dim nVidiaSMI, OutputFormat, oShell, oFSO, sTempFile, Return, oFile, oGPUPowerLevels, GPUPowerLevels, aGPUPowerLevels, PowerLimit, sLogFile, fLogFile, i, Total

nVidiaSMI="""C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"""
OutputFormat="--format=csv,noheader,nounits"

Set oShell=CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")
Const ForAppending = 8
Const CreateIfNotExist = 1
Const OpenAsASCII = 0
sTempFile=oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName
sLogFile = "C:\Scripts\Logs\Monitor-PowerLevels.log"
PowerLimit = "120.00"

Return = oShell.Run("cmd /c " & nVidiaSMI & " --query-gpu=power.limit " & OutputFormat & " > " & sTempFile, 0, True)
Set oFile = oFSO.OpenTextFile(sTempFile, 1)
GPUPowerLevels = oFile.ReadAll
Trim(GPUPowerLevels)
oFile.Close
oFSO.DeleteFile(sTempFile)

aGPUPowerLevels=Split(GPUPowerLevels,vbCrLf)
For i = 0 To UBound(aGPUPowerLevels)-1
	aGPUPowerLevels(i) = Trim(aGPUPowerLevels(i))
	If aGPUPowerLevels(i) = PowerLimit Then
		'WScript.Echo "GPU#" & i & " Power limit " & aGPUPowerLevels(i) & " good"
	Else
		Set fLogFile = oFSO.OpenTextFile(sLogFile, ForAppending, CreateIfNotExist, OpenAsASCII)
		fLogFile.WriteLine ("GPU#" & i & " Power limit " & aGPUPowerLevels(i) & " bad, changing to " & PowerLimit & " @ " & Now())
		fLogFile.Close
		Return = oShell.Run("cmd /c " & nVidiaSMI & " -i " & i & " -pl " & PowerLimit, 0, True)		
	End If
Next
