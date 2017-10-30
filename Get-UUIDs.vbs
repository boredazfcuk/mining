Option Explicit

Dim nVidiaSMI, QueryUUIDs, OutputFormat, oShell, oFSO, sTempFile, Return, oFile, GPUUUIDs, aGPUUUIDs, i

nVidiaSMI="""C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"""
QueryUUIDs=" --query-gpu=uuid "
OutputFormat="--format=csv,noheader,nounits"

Set oShell=CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")
Const ForAppending = 8
Const CreateIfNotExist = 1
Const OpenAsASCII = 0

' Check if running under "WScript", and if so, relaunch in cscript
If InStr(1, WScript.FullName, "WScript.exe", vbTextCompare) <> 0 Then
        oShell.Run "%comspec% /c cscript /nologo """ & WScript.ScriptFullName & """", 1, False
        WScript.Quit(0)
End If

' Generate temporary file name
sTempFile=oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName

' Perform GPU UUID Query and purh results to the temp file
Return = oShell.Run("cmd /c " & nVidiaSMI & QueryUUIDs & OutputFormat & " > " & sTempFile, 0, True)
' Open temp file for reading
Set oFile = oFSO.OpenTextFile(sTempFile, 1)
' Read whole file into a variable
GPUUUIDs = oFile.ReadAll
' Trim the variable to remove trailing whitespace etc
Trim(GPUUUIDs)
' Close the temp file
oFile.Close
' Delete the temp file
oFSO.DeleteFile(sTempFile)

' Split the query results into an array
aGPUUUIDs=Split(GPUUUIDs,VBNewLine)
' For each variable in the array except the last, empty one
For i = LBound(aGPUUUIDs) to UBound(aGPUUUIDs)-1
	' Print the GPU UUID to screen
	WScript.Echo aGPUUUIDs(i)
Next

' Kick out a prompt to user
WScript.Echo "Press [ENTER] to continue..."
' Wait until [ENTER] is pressed.
WScript.StdIn.ReadLine
