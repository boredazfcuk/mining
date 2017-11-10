Option Explicit

'
' THIS SCRIPT IS PROVIDED "AS IS", USE AT YOUR OWN RISK!
' https://github.com/boredazfcuk/mining
'

'----- Initialise Variables -----
Dim oFSO, oShell, oRegistry, nVidiaSMI, QueryGPUCount, QueryGPUUUIDs, QuerySensorValues, OutputFormat, RegKey, KeyPath, ValueName, GPUKeyExists
Dim sTempFile, TargetGPU, RunSilent, oTempFile, Count, SensorResults, aSensorResults
Dim GPUUtilisation, GPUUtilisationFree, GPUTemperature, GPUFanSpeed, GPUMemTotal, GPUMemUsed, GPUPowerDraw

Const HKLM = &H80000002

'----- Exit if GPU Index number not supplied -----
If WScript.Arguments.Count <> 1 Then
	'----- Kick out error message -----
	WScript.Echo "Usage: cscript.exe //nologo GPUDetails.vbs <GPUNumber>"
	'----- Exit out with error message -----
	WScript.Quit(1)
End If

'----- Create Objects -----
Set oFSO=CreateObject("Scripting.FileSystemObject")
Set oShell=CreateObject("WScript.Shell")
Set oRegistry=GetObject("winmgmts:\\.\root\default:StdRegProv")

'----- Set Variables -----
nVidiaSMI="""C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"""
QueryGPUCount=" -i 0 --query-gpu=count "
QueryGPUUUIDs=" --query-gpu=uuid "
QuerySensorValues=" --query-gpu=utilization.gpu,temperature.gpu,fan.speed,memory.total,memory.used,power.draw "
OutputFormat="--format=csv,noheader,nounits"
RegKey="HKLM\Software\Wow6432Node\boredazfcuk\mining\GPUs\"
KeyPath="Software\WOW6432Node\boredazfcuk\mining\GPUs\"

'----- Get Temp File -----
sTempFile=oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName

'----- Read GPU UUID from registry -----
oRegistry.GetStringValue HKLM, KeyPath, WScript.Arguments.Item(0), GPUKeyExists
If IsNull (GPUKeyExists) Then
	GetGPUUUIDs
End If
TargetGPU=oShell.RegRead(RegKey & WScript.Arguments.Item(0))


'----- Query GPU 3 times in ~10 seconds -----
QueryGPU
WScript.Sleep(5000)
QueryGPU
WScript.Sleep(5000)
QueryGPU

'----- If GPU has errored -----
If GPUMemTotal="[Unknown Error]" Then
	'----- Set amount of memory used to 1KB -----
	GPUMemUsed=1
Else
	'----- Convert results to Bytes from MB -----
	GPUMemTotal=GPUMemTotal*1024*1024
	GPUMemUsed=GPUMemUsed*1024*1024
End If

'----- Kick out XML result for PRTG Network Monitor -----
WScript.Echo "<?xml version=""1.0"" encoding=""Windows-1252""?>"
WScript.Echo "<PRTG>"
WScript.Echo "	<result>"
WScript.Echo "		<channel>GPU Utilization</channel>"
WScript.Echo "		<unit>Percentage</unit>"
WScript.Echo "		<mode>Absolute</mode>"
WScript.Echo "		<showChart>1</showChart>"
WScript.Echo "		<showTable>1</showTable>"
WScript.Echo "		<value>" & Round(GPUUtilisation/3) & "</value>"
WScript.Echo "		<warning>0</warning>"
WScript.Echo "		<float>0</float>"
WScript.Echo "		<LimitMinError>70</LimitMinError>"
WScript.Echo "		<LimitMinWarning>90</LimitMinWarning>"
WScript.Echo "		<LimitWarningMsg>Mining Performance Impacted</LimitWarningMsg>"
WScript.Echo "		<LimitErrorMsg>Mining Failure</LimitErrorMsg>"
WScript.Echo "		<LimitMode>1</LimitMode>"
WScript.Echo "	</result>"
WScript.Echo "	<result>"
WScript.Echo "		<channel>Temperature</channel>"
WScript.Echo "		<unit>Temperature</unit>"
WScript.Echo "		<mode>Absolute</mode>"
WScript.Echo "		<showChart>1</showChart>"
WScript.Echo "		<showTable>1</showTable>"
WScript.Echo "		<warning>0</warning>"
WScript.Echo "		<value>" & Round(GPUTemperature/3) & "</value>"
WScript.Echo "		<float>0</float>"
WScript.Echo "		<LimitMaxError>90</LimitMaxError>"
WScript.Echo "		<LimitMaxWarning>70</LimitMaxWarning>"
WScript.Echo "		<LimitWarningMsg>Temperature High</LimitWarningMsg>"
WScript.Echo "		<LimitErrorMsg>Temperature Danger</LimitErrorMsg>"
WScript.Echo "		<LimitMode>1</LimitMode>"
WScript.Echo "	</result>"
WScript.Echo "	<result>"
WScript.Echo "		<channel>Fan Speed</channel>"
WScript.Echo "		<unit>Percentage</unit>"
WScript.Echo "		<mode>Absolute</mode>"
WScript.Echo "		<showChart>1</showChart>"
WScript.Echo "		<showTable>1</showTable>"
WScript.Echo "		<value>" & Round(GPUFanSpeed/3) & "</value>"
WScript.Echo "		<warning>0</warning>"
WScript.Echo "		<float>0</float>"
WScript.Echo "		<LimitMode>0</LimitMode>"
WScript.Echo "	</result>"
WScript.Echo "	<result>"
WScript.Echo "		<channel>Power Draw</channel>"
WScript.Echo "		<unit>Count</unit>"
WScript.Echo "		<mode>Absolute</mode>"
WScript.Echo "		<showChart>1</showChart>"
WScript.Echo "		<showTable>1</showTable>"
WScript.Echo "		<warning>0</warning>"
WScript.Echo "		<value>" & Round(GPUPowerDraw/3) & "</value>"
WScript.Echo "		<float>1</float>"
WScript.Echo "		<LimitMaxError>125</LimitMaxError>"
WScript.Echo "		<LimitMinError>60</LimitMinError>"
WScript.Echo "		<LimitErrorMsg>Undervolt Error</LimitErrorMsg>"
WScript.Echo "		<LimitMode>1</LimitMode>"
WScript.Echo "	</result>"
WScript.Echo "	<result>"
WScript.Echo "		<channel>Available Memory</channel>"
WScript.Echo "		<unit>BytesMemory</unit>"
WScript.Echo "		<mode>Absolute</mode>"
WScript.Echo "		<showChart>1</showChart>"
WScript.Echo "		<showTable>1</showTable>"
WScript.Echo "		<value>" & Round(GPUMemUsed/3) & "</value>"
WScript.Echo "		<warning>0</warning>"
WScript.Echo "		<float>0</float>"
WScript.Echo "		<LimitMinError>" & 1024*1024*128 & "</LimitMinError>"
WScript.Echo "		<LimitMinWarning>" & 1024*1024*256 & "</LimitMinWarning>"
WScript.Echo "		<LimitWarningMsg>Memory Low</LimitWarningMsg>"
WScript.Echo "		<LimitErrorMsg>Memory Critical</LimitErrorMsg>"
WScript.Echo "		<LimitMode>1</LimitMode>"
WScript.Echo "	</result>"
WScript.Echo "	<result>"
WScript.Echo "		<channel>GPU#</channel>"
WScript.Echo "		<value>" & WScript.Arguments.Item(0) & "</value>"
WScript.Echo "		<unit>Count</unit>"
WScript.Echo "		<text>UUID=" & TargetGPU & "</text>"
WScript.Echo "	</result>"
WScript.Echo "</PRTG>"

Sub QueryGPU
	'----- Query Values to be returned by sensor -----
	RunSilent=oShell.Run("cmd /c " & nVidiaSMI & " -i " & TargetGPU & QuerySensorValues & OutputFormat & " > " & sTempFile, 0, True)
	
	'----- Target Temp File -----
	Set oTempFile=oFSO.OpenTextFile(sTempFile, 1)
	
	'----- Loop through Results -----
	Do While Not oTempFile.AtEndOfStream
		Count=0
		'----- Read line -----
		SensorResults=oTempFile.ReadLine()
		'----- Clean up Line -----
		Trim(SensorResults)
		'----- Split line into array using comma as separator -----
		aSensorResults=Split(SensorResults, ",")
		'----- Check each array element -----
		For Count=0 To UBound(aSensorResults)
			'----- Clean up value -----
			aSensorResults(Count)=Trim(aSensorResults(Count))
		Next
	Loop
	'----- Put results into individual variables, adding to previous values -----
	GPUUtilisation=Int(GPUUtilisation)+Int(aSensorResults(0))
	GPUTemperature=Int(GPUTemperature)+Int(aSensorResults(1))
	GPUFanSpeed=Int(GPUFanSpeed)+Int(aSensorResults(2))
	GPUMemUsed=Int(GPUMemUsed)+Int(aSensorResults(4))
	GPUPowerDraw=Int(GPUPowerDraw)+Int(aSensorResults(5))
End Sub

Sub GetGPUUUIDs
	'----- Initilise variables -----
	Dim oFile, GPUUUIDs, aGPUUUIDs, Count
	
	'----- Perform GPU UUID Query and push results to the temp file -----
	RunSilent = oShell.Run("cmd /c " & nVidiaSMI & QueryGPUUUIDs & OutputFormat & " > " & sTempFile, 0, True)
	'----- Open temp file for reading -----
	Set oFile = oFSO.OpenTextFile(sTempFile, 1)
	'----- Read whole file into a variable -----
	GPUUUIDs = oFile.ReadAll
	'----- Trim the variable to remove trailing whitespace etc -----
	Trim(GPUUUIDs)
	'----- Close the temp file -----
	oFile.Close
	'----- Delete the temp file -----
	oFSO.DeleteFile(sTempFile)
	
	'----- Split the query results into an array -----
	aGPUUUIDs=Split(GPUUUIDs,VBNewLine)
	'----- For each variable in the array except the last, empty one -----
	For Count = LBound(aGPUUUIDs) to UBound(aGPUUUIDs)-1
		'----- Save to registry -----
		oShell.RegWrite RegKey & Count, aGPUUUIDs(Count), "REG_SZ"
	Next
End Sub

'----- Close Temp file -----
oTempFile.Close
'----- Delete Temp file -----
oFSO.DeleteFile(sTempFile)