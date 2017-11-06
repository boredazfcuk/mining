Option Explicit
'
' THIS SCRIPT IS PROVIDED "AS IS", USE AT YOUR OWN RISK!
' https://github.com/boredazfcuk/mining
'
'-----
' This script requires two files to run. These should both be located in the root of your user profile
'
' The first (EncryptionKey.txt) is a plain text file that contains the encryption key with which to encrypt
' the archive file that is created by 7-zip.
'
' The second is a csv file that contains a reference to the currency you are creating. This doubles as the
' file name used to store the backup. It also contains the path to back the file up to. It also contains
' a pointer to the wallet file for that currency. The last is a pointer to a folder that you would also
' like backed up. Here is an example of how it needs to be formatted:
'
'	Bitcoin,\OwnCloud,\AppData\Roaming\BitcoinWalletApp\wallet.dat,\AppData\Roaming\BitcoinWalletApp\backups
'	Ethereum,\OwnCloud,\AppData\Roaming\EthWalltApp\eth.wallet,\AppData\Roaming\EthWalltApp\Partitions
'
'-----

'----- Initialise Variables -----
Dim SevenZip, oShell, ProfilePath, EncryptionKeyFile, BackupConfigFile, CurrentFolder, LogFolder
Dim oFSO, sLogFile, fLogFile, fEncryptionKeyFile, fBackupConfigFile, EncryptionKey, BackupConfig
Dim Line, BackupPath, Count, aBackupConfig, CryptoCurrency, WalletFile, BackupTarget, RunSilent

'----- Set Constants -----
Const ForAppending = 8
Const ForReading = 1
Const CreateIfNotExist = 1
Const OpenAsASCII = 0

'----- Set 7-zip path -----
SevenZip="""C:\Program Files\7-Zip\7z.exe"""

'----- Create Objects -----
Set oShell = WScript.CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")

'----- Set Variables -----
ProfilePath=oShell.ExpandEnvironmentStrings("%UserProfile%")
EncryptionKeyFile=ProfilePath & "\EncryptionKey.txt"
BackupConfigFile=ProfilePath & "\BackupConfig.csv"

'----- Get Script Folder -----
CurrentFolder = oFSO.GetAbsolutePathName(".")
LogFolder = oFSO.BuildPath(CurrentFolder, "\Logs")
'----- If Log Sub Folder doesn't exist -----
If Not (oFSO.FolderExists(LogFolder)) Then
    '----- Create Log SubFolder-----
    oFSO.CreateFolder(LogFolder)
End If
sLogFile = oFSO.BuildPath(LogFolder, "\Backup-Wallets.log")

'----- Target Log and Config Files -----
Set fLogFile = oFSO.OpenTextFile(sLogFile, ForAppending, CreateIfNotExist, OpenAsASCII)
Set fEncryptionKeyFile = oFSO.OpenTextFile(EncryptionKeyFile, ForReading, CreateIfNotExist, OpenAsASCII)
Set fBackupConfigFile = oFSO.OpenTextFile(BackupConfigFile, ForReading, CreateIfNotExist, OpenAsASCII)

'----- Read Encryption Key from File -----
EncryptionKey=fEncryptionKeyFile.ReadLine()

'----- Loop through backup config file performing backups -----
Do While Not fBackupConfigFile.AtEndOfStream
	Count=0
	'----- Read line -----
	BackupConfig = fBackupConfigFile.ReadLine()
	'----- Split line into array using comma as separator -----
	aBackupConfig = Split(BackupConfig, ",")
	'----- Check each array element -----
	For Count = 0 To UBound(aBackupConfig)
		'----- Clean up value -----
		aBackupConfig(Count) = Trim(aBackupConfig(Count))
	Next
	'----- Set Column one value to Cryptocurrency type -----
	CryptoCurrency=aBackupConfig(0)
	'----- Set Column two value to backup location -----
	BackupPath=aBackupConfig(1)
	'----- Set Column three value to be wallet file to add to archive -----
	WalletFile=aBackupConfig(2)
	'----- Set Column four value to be wallet folder to add to archive -----
	BackupTarget=aBackupConfig(3)

	'----- Perform backup -----
	RunSilent=oShell.Run (SevenZip & " a -p" & EncryptionKey & " -mx9 " & ProfilePath & BackupPath & "\" & CryptoCurrency & ".7z " & ProfilePath & WalletFile & " " & ProfilePath & BackupTarget, 0, False)
	'----- Write Event to Windows Application Event Viewer Log -----
	oShell.LogEvent 0, CryptoCurrency & " wallet backed up to " & ProfilePath & BackupPath & "\" & CryptoCurrency & ".7z"
	'----- Write entry to log file -----
	fLogFile.WriteLine (CryptoCurrency & " wallet backed up to " & ProfilePath & BackupPath & "\" & CryptoCurrency & ".7z @ " & Now())
Loop
'----- Close Files -----
fLogFile.Close
fEncryptionKeyFile.Close
fBackupConfigFile.Close