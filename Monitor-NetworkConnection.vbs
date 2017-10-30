Option Explicit

Dim oShell, oFSO, sLogFile, fLogFile, RouterIP, PingAttempts
Set oShell=CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")
Const OpenAsASCII = 0 
Const FailIfNotExist = 0
Const CreateIfNotExist = 1
Const ForReading =  1
Const ForAppending = 8
sLogFile = "C:\Scripts\Logs\Monitor-NetworkConnection.log"
RouterIP="192.168.1.254"
PingAttempts = 1

If IsAlive(RouterIP) Then
	'WScript.Echo("He's Alive! ALIVE!")
Else
	PingAttempts = 60
	If IsAlive(RouterIP) Then
		'WScript.Echo("Little wobble... but things are OK now")
	Else
		Set fLogFile = oFSO.OpenTextFile(sLogFile, ForAppending, CreateIfNotExist, OpenAsASCII)
		fLogFile.WriteLine ("Network connection lost at " & Now() & " - Rebooting.")
		fLogFile.Close
		Set oShell=CreateObject("WScript.Shell")
		oShell.Run "%comspec% /c shutdown /f /r /t 60", , True
	End If
End If

Function IsAlive(sHost) 
    Dim sTempFile, fFile  
    sTempFile = oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName 
    oShell.Run "%comspec% /c ping.exe -n " & PingAttempts & " " & sHost & ">" & sTempFile, 0 , True 
    Set fFile = oFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, OpenAsASCII) 
    Select Case InStr(fFile.ReadAll, "TTL=") 
         Case 0
            IsAlive = False 
         Case Else
            IsAlive = True 
    End Select 
    fFile.Close 
    oFSO.DeleteFile(sTempFile)
End Function