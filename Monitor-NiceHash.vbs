Option Explicit

Const processName   = "NiceHashMinerLegacy.exe"
Const maxAgeSeconds = 180
Const OpenAsASCII = 0 
Const CreateIfNotExist = 1
Const ForAppending = 8

Dim oWMI, objItem, colItems, oFSO, sTempFile, oFile, Return, sLogFile, fLogFile, UtilisationTotal

Set oWMI = GetObject("winmgmts:\\localhost\root\CIMV2")
Set colItems = oWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Caption='" & processName & "'")
Set oFSO  = CreateObject("Scripting.FileSystemObject")
sTempFile = oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName 
sLogFile = "C:\Scripts\Logs\Monitor-NiceHash.log"
UtilisationTotal = 0

For Each objItem In colItems
    Dim secAge
    secAge = DateDiff("s", WMIDateStringToDate(objItem.CreationDate), Now())
   
    If secAge > maxAgeSeconds Then
        CheckUtilisation
        WScript.Sleep(3000)
        CheckUtilisation
        WScript.Sleep(3000)
        CheckUtilisation
        WScript.Sleep(3000)
        CheckUtilisation
        If UtilisationTotal = 4 Then
        
        	Dim sNiceHash, sExcavator, sXMRStakCPU, ssgminer, snheqminer, sethminer, sccminer, sprospector, soptiminer, sminer, szecminer64, sethdcrminer64, snsgpucnminer, oProcess, cProcess, oShell

			sNiceHash = "'NiceHashMinerLegacy.exe'" 
			sExcavator = "'Excavator.exe'"
			sXMRStakCPU = "'xmr-stak-cpu.exe'"
			ssgminer = "'sgminer.exe'"
			snheqminer = "'nheqminer.exe'"
			sethminer = "'ethminer.exe'"
			sccminer = "'ccminer.exe'"
			sprospector = "'prospector.exe'"
			soptiminer = "'optiminer.exe'"
			sminer = "'miner.exe'"
			szecminer64 = "'zecminer64.exe'"
			sethdcrminer64 = "'ethdcrminer64.exe'"
			snsgpucnminer = "'nsgpucnminer.exe'"
			
        	Set oShell=CreateObject("WScript.Shell")
			Set fLogFile = oFSO.OpenTextFile(sLogFile, ForAppending, CreateIfNotExist, OpenAsASCII)
			fLogFile.WriteLine ("GPU Utilisation below 80% at " & Now() & " - Restarting Nice Hash.")
			fLogFile.Close
			Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 
			Set cProcess = oWMI.ExecQuery ("Select * from Win32_Process WHERE Name = " & sNiceHash & " OR Name = " & sExcavator & " OR Name = " & sXMRStakCPU & " OR Name = " & ssgminer & " OR Name = " & snheqminer & " OR Name = " & sethminer & " OR Name = " & sccminer & " OR Name = " & sprospector & " OR Name = " & soptiminer & " OR Name = " & sminer & " OR Name = " & szecminer64 & " OR Name = " & sethdcrminer64 & " OR Name = " & snsgpucnminer & " OR Name = " & sethminer)
			For Each oProcess in cProcess
				oProcess.Terminate()
			Next
			WScript.Sleep(3000)
			oShell.Run("""C:\Program Files\NiceHash Miner Legacy\NiceHashMinerLegacy.exe""")
		End If
    End If
Next

Sub CheckUtilisation

	Dim nVidiaSMI, OutputFormat, oShell, GPUDevices, GPUUtilisation, i, aGPUUtilisation
	Dim Total, UtilisationAverage, oProcess, cProcess
	Dim sNiceHash, sExcavator, sXMRStakCPU, ssgminer, snheqminer, sethminer, sccminer, sprospector, soptiminer, sminer, szecminer64, sethdcrminer64, snsgpucnminer

	sNiceHash = "'NiceHashMinerLegacy.exe'" 
	sExcavator = "'Excavator.exe'"
	sXMRStakCPU = "'xmr-stak-cpu.exe'"
	ssgminer = "'sgminer.exe'"
	snheqminer = "'nheqminer.exe'"
	sethminer = "'ethminer.exe'"
	sccminer = "'ccminer.exe'"
	sprospector = "'prospector.exe'"
	soptiminer = "'optiminer.exe'"
	sminer = "'miner.exe'"
	szecminer64 = "'zecminer64.exe'"
	sethdcrminer64 = "'ethdcrminer64.exe'"
	snsgpucnminer = "'nsgpucnminer.exe'"

	Total=0

	nVidiaSMI="""C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"""
	OutputFormat="--format=csv,noheader,nounits"

	Set oShell=CreateObject("WScript.Shell")
	Return = oShell.Run("cmd /c " & nVidiaSMI & " -i 0 --query-gpu=count " & OutputFormat & " > " & sTempFile, 0, True)
	Set oFile = oFSO.OpenTextFile(sTempFile, 1)
	GPUDevices = oFile.ReadLine
	Trim(GPUDevices)
	oFile.Close
	
	Return = oShell.Run("cmd /c " & nVidiaSMI & " --query-gpu=utilization.gpu " & OutputFormat & " > " & sTempFile, 0, True)
	
	Set oFile = oFSO.OpenTextFile(sTempFile, 1)
	GPUUtilisation = oFile.ReadAll
	Trim(GPUDevices)
	oFile.Close
	oFSO.DeleteFile(sTempFile)
	
	aGPUUtilisation=Split(GPUUtilisation,vbCrLf)
	For i = 0 To UBound(aGPUUtilisation)-1
		aGPUUtilisation(i) = Trim(aGPUUtilisation(i))
		Total = Total + Int(aGPUUtilisation(i))
	Next
	UtilisationAverage= Total / UBound(aGPUUtilisation)
	If UtilisationAverage < 80 Then
		UtilisationTotal= UtilisationTotal+1
	End If
End Sub

Function WMIDateStringToDate(dtmDate)
     WMIDateStringToDate = CDate(Mid(dtmDate, 7, 2) & "/" & _
     Mid(dtmDate, 5, 2) & "/" & Left(dtmDate, 4) _
     & " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function