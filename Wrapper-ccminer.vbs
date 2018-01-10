Option Explicit

'
' THIS SCRIPT IS PROVIDED "AS IS", USE AT YOUR OWN RISK!
' https://github.com/boredazfcuk/mining
'

'----- Create Variables -----
Dim Miner, Parameters, oShell, sScriptName, oFSO, oFile, ScriptFolder, Return

Miner="ccminer.exe"
Parameters="-a xxx -o stratum+tcp://xxx.xxx.xxx:3739 -u XxxXxxXxXX -p c=XXX -i 20"

'----- Create Shell Object ------
Set oShell=WScript.CreateObject("WScript.Shell")

'----- Get Script Name -----
sScriptName=WScript.ScriptFullName
'----- Get Script Folder -----
Set oFSO=CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.GetFile(sScriptName)
ScriptFolder=oFSO.GetParentFolderName(oFile)

'----- Keep checking the process trace -----
Do
	'----- Launch Miner -----
	Return=oShell.Run(ScriptFolder & Chr(92) & Miner & Chr(32) & Parameters, 0, True)
Loop