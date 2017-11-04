Option Explicit

Dim oShell, TargetGPU, GPUDevices, oGPUDevices, GPUUUIDs, oGPUUUIDs, aGPUUUIDS, i, GPUIndex, GPUResults
Dim oGPUResults, aGPUResults, GPUUUID, GPUName, GPUBIOS, GPUDriver, GPUUtilization, GPUUtilisationFree, nVidiaSMI
Dim OutputFormat, GPUTemperature, GPUFanSpeed, GPUMemTotal, GPUMemUsed, GPUMemFree, GPUPowerDraw, RegKey

nVidiaSMI="""C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"""
OutputFormat="--format=csv,noheader,nounits"
RegKey="HKLM\Software\Wow6432Node\boredazfcuk\mining\GPUs\"

Set oShell=CreateObject("WScript.Shell")
Set oGPUDevices = oShell.Exec(nVidiaSMI & " -i 0 --query-gpu=count " & OutputFormat)
Set oGPUUUIDs = oShell.Exec(nVidiaSMI & " --query-gpu=uuid " & OutputFormat)

If WScript.Arguments.Count <> 1 then
	WScript.Echo "Usage: cscript.exe //nologo GPUDetails.vbs <GPUNumber>"
	WScript.Quit(0)
End If
TargetGPU=oShell.RegRead(RegKey & WScript.Arguments.Item(0))
Trim(TargetGPU)

Do While Not oGPUDevices.StdOut.AtEndOfStream
	GPUDevices = oGPUDevices.StdOut.ReadLine
	Trim(GPUDevices)
Loop

Do While Not oGPUUUIDs.StdOut.AtEndOfStream
	GPUUUIDs = oGPUUUIDs.StdOut.ReadAll
Loop
aGPUUUIDs=Split(GPUUUIDs,VBNewLine)
For i = LBound(aGPUUUIDs) to UBound(aGPUUUIDs)
	If aGPUUUIDs(i) = TargetGPU Then
		GPUIndex = i
	End If
Next

Set oGPUResults = oShell.Exec(nVidiaSMI & " -i " & GPUIndex & " --query-gpu=uuid,name,vbios_version,driver_version,utilization.gpu,temperature.gpu,fan.speed,memory.total,memory.used,memory.free,power.draw " & OutputFormat)

Do While Not oGPUResults.StdOut.AtEndOfStream
	GPUResults = oGPUResults.StdOut.ReadAll
Loop

aGPUResults=Split(GPUResults,",")
For i = 0 To UBound(aGPUResults)
	aGPUResults(i) = Trim(Replace(aGPUResults(i), vbCrLf, ""))
Next

GPUUUID = aGPUResults(0)
GPUName = aGPUResults(1)
GPUBIOS = aGPUResults(2)
GPUDriver = aGPUResults(3)
GPUUtilization = aGPUResults(4)
GPUUtilisationFree = 100-GPUUtilization
GPUTemperature = aGPUResults(5)
GPUFanSpeed = aGPUResults(6)
GPUMemTotal = aGPUResults(7)
GPUMemUsed = aGPUResults(8)
GPUMemFree = aGPUResults(9)
GPUPowerDraw = aGPUResults(10)
If GPUMemTotal = "[Unknown Error]" Then
	GPUMemUsed = 0
Else
	GPUMemTotal = GPUMemTotal*1024*1024
	GPUMemUsed = GPUMemUsed*1024*1024
	GPUMemFree =GPUMemFree*1024*1024
End If
 
wscript.echo "<?xml version=""1.0"" encoding=""Windows-1252""?>"
wscript.echo "<PRTG>"
wscript.echo "	<result>"
wscript.echo "		<channel>GPU Utilization</channel>"
wscript.echo "		<unit>Percentage</unit>"
wscript.echo "		<mode>Absolute</mode>"
wscript.echo "		<showChart>1</showChart>"
wscript.echo "		<showTable>1</showTable>"
wscript.echo "		<value>" & GPUUtilization & "</value>"
wscript.echo "		<warning>0</warning>"
wscript.echo "		<float>0</float>"
wscript.echo "		<LimitMinError>80</LimitMinError>"
wscript.echo "		<LimitMinWarning>90</LimitMinWarning>"
wscript.echo "		<LimitWarningMsg>Mining Performance Impacted</LimitWarningMsg>"
wscript.echo "		<LimitErrorMsg>Mining Failure</LimitErrorMsg>"
wscript.echo "		<LimitMode>1</LimitMode>"
wscript.echo "	</result>"
wscript.echo "	<result>"
wscript.echo "		<channel>Temperature</channel>"
wscript.echo "		<unit>Temperature</unit>"
wscript.echo "		<mode>Absolute</mode>"
wscript.echo "		<showChart>1</showChart>"
wscript.echo "		<showTable>1</showTable>"
wscript.echo "		<warning>0</warning>"
wscript.echo "		<value>" & GPUTemperature & "</value>"
wscript.echo "		<float>0</float>"
wscript.echo "		<LimitMaxError>90</LimitMaxError>"
wscript.echo "		<LimitMaxWarning>70</LimitMaxWarning>"
wscript.echo "		<LimitWarningMsg>Temperature High</LimitWarningMsg>"
wscript.echo "		<LimitErrorMsg>Temperature Danger</LimitErrorMsg>"
wscript.echo "		<LimitMode>1</LimitMode>"
wscript.echo "	</result>"
wscript.echo "	<result>"
wscript.echo "		<channel>Fan Speed</channel>"
wscript.echo "		<unit>Percentage</unit>"
wscript.echo "		<mode>Absolute</mode>"
wscript.echo "		<showChart>1</showChart>"
wscript.echo "		<showTable>1</showTable>"
wscript.echo "		<value>" & GPUFanSpeed & "</value>"
wscript.echo "		<warning>0</warning>"
wscript.echo "		<float>0</float>"
wscript.echo "		<LimitMaxError>90</LimitMaxError>"
wscript.echo "		<LimitMaxWarning>70</LimitMaxWarning>"
wscript.echo "		<LimitWarningMsg>Fan Speed High</LimitWarningMsg>"
wscript.echo "		<LimitErrorMsg>Fan Speed Critical</LimitErrorMsg>"
wscript.echo "		<LimitMode>1</LimitMode>"
wscript.echo "	</result>"
wscript.echo "	<result>"
wscript.echo "		<channel>Power Draw</channel>"
wscript.echo "		<unit>Count</unit>"
wscript.echo "		<mode>Absolute</mode>"
wscript.echo "		<showChart>1</showChart>"
wscript.echo "		<showTable>1</showTable>"
wscript.echo "		<warning>0</warning>"
wscript.echo "		<value>" & GPUPowerDraw & "</value>"
wscript.echo "		<float>1</float>"
wscript.echo "		<LimitMaxError>125</LimitMaxError>"
wscript.echo "		<LimitMinError>60</LimitMinError>"
wscript.echo "		<LimitErrorMsg>Undervolt Error</LimitErrorMsg>"
wscript.echo "		<LimitMode>1</LimitMode>"
wscript.echo "	</result>"
wscript.echo "	<result>"
wscript.echo "		<channel>Available Memory</channel>"
wscript.echo "		<unit>BytesMemory</unit>"
wscript.echo "		<mode>Absolute</mode>"
wscript.echo "		<showChart>1</showChart>"
wscript.echo "		<showTable>1</showTable>"
wscript.echo "		<value>" & GPUMemUsed & "</value>"
wscript.echo "		<warning>0</warning>"
wscript.echo "		<float>0</float>"
wscript.echo "		<LimitMinError>" & 1024*1024*128 & "</LimitMinError>"
wscript.echo "		<LimitMinWarning>" & 1024*1024*256 & "</LimitMinWarning>"
wscript.echo "		<LimitWarningMsg>Memory Low</LimitWarningMsg>"
wscript.echo "		<LimitErrorMsg>Memory Critical</LimitErrorMsg>"
wscript.echo "		<LimitMode>1</LimitMode>"
wscript.echo "	</result>"
wscript.echo "	<result>"
wscript.echo "		<channel>GPU#</channel>"
wscript.echo "		<value>" & GPUIndex & "</value>"
wscript.echo "		<unit>Count</unit>"
wscript.echo "		<text>UUID=" & TargetGPU & "</text>"
wscript.echo "	</result>"
wscript.echo "</PRTG>"