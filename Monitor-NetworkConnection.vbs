Option Explicit

'
' https://github.com/boredazfcuk/mining
'

'----- Create Variables -----
Dim oShell, oFSO, CurrentFolder, LogFolder, sLogFile, oWMI, cGateways, Gateway, DefaultGateway, PingAttempts

'----- Create Objects -----
Set oShell=CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")
Set oWMI=GetObject("winmgmts:\\localhost\root\CIMV2")

'----- Set Constants -----
Const OpenAsASCII = 0 
Const FailIfNotExist = 0
Const CreateIfNotExist = 1
Const ForReading =  1
Const ForAppending = 8

'----- Get Script Folder -----
CurrentFolder = oFSO.GetAbsolutePathName(".")
LogFolder = oFSO.BuildPath(CurrentFolder, "\Logs")
'----- If Log Sub Folder doesn't exist -----
If Not (oFSO.FolderExists(LogFolder)) Then
    '----- Create Log SubFolder-----
    oFSO.CreateFolder(LogFolder)
End If
'----- Create full log file path -----
sLogFile = oFSO.BuildPath(LogFolder, "\Monitor-NetworkConnection.log")

'----- Fetch gateways for all IP enabled NICs from WMI -----
Set cGateways = oWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=TRUE")
'----- For each gateway in the gateway array -----
For Each Gateway In cGateways
		'----- Grab the first gateway and quit out, fingers crossed this isn't a multi-homed client -----
		DefaultGateway=Gateway.DefaultIPGateway(0)
	Exit For
Next

'----- Set the number of pings for the initial test ---
PingAttempts = 1

'----- Check if the Gateway is responding with a single ping -----
If IsAlive(DefaultGateway) Then
	'----- All is good so do nothing -----
	'WScript.Echo("He's Alive! ALIVE!")
Else
	'----- If the ping failed, increase the number of attempts to 30 and see if all of them are lost -----
	PingAttempts = 10
	'----- Prolonged check to see if the gateway is responding -----
	If IsAlive(DefaultGateway) Then
		'----- Original ping was lost but at least one got through -----
		'WScript.Echo("Little wobble... but things are OK now")
	Else
		'----- All pings were lost so log the result and reboot the computer -----
		Set fLogFile = oFSO.OpenTextFile(sLogFile, ForAppending, CreateIfNotExist, OpenAsASCII)
		'----- Log to the event log -----
		oShell.LogEvent 1, "Network connection lost at " & Now() & " - Rebooting."
		'----- Log to the log file too ----
		fLogFile.WriteLine ("Network connection lost at " & Now() & " - Rebooting.")
		'----- Close the log file -----
		fLogFile.Close
		Set oShell=CreateObject("WScript.Shell")
		'----- Reboot the computer -----
		oShell.Run "%comspec% /c shutdown /f /r /t 60", , True
	End If
End If

'----- Function to ping the default gateway address -----
Function IsAlive(sHost) 
    Dim sTempFile, fFile  
    '----- Get a name for the Temp file -----
    sTempFile = oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName 
    '----- Run the ping command and log the results to a Temp file -----
    oShell.Run "%comspec% /c ping.exe -n " & PingAttempts & " " & sHost & ">" & sTempFile, 0 , True 
    '----- Prepare to write to the log file -----
    Set fFile = oFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, OpenAsASCII) 
    '----- Check to see if there was a reply -----
    Select Case InStr(fFile.ReadAll, "TTL=")
    	'----- If not alive, set flag -----
         Case 0
            IsAlive = False 
        '----- If is alive, set flag -----
         Case Else
            IsAlive = True 
    End Select
    '----- Close text file -----
    fFile.Close
    '----- Delete Temp file -----
    oFSO.DeleteFile(sTempFile)
End Function