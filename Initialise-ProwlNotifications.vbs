Option Explicit

Dim oShell, oRegistry, sRegKey, sKeyPath, sValueName, ProwlAPIKey,  ProwlKeyExists

Const HKCU = &H80000001

Set oShell=CreateObject("WScript.Shell")
Set oRegistry=GetObject("winmgmts:\\.\root\default:StdRegProv")

sRegKey="HKCU\Software\boredazfcuk\mining\ProwlAPIKey"
sKeyPath="Software\boredazfcuk\mining"
sValueName="ProwlAPIKey"
ProwlAPIKey=InputBox("Please enter your Prowl API Key (Cancel to delete):", "Prowl API")

If ((ProwlAPIKey=Null) Or (ProwlAPIKey="")) Then
	oRegistry.GetStringValue HKCU, sKeyPath, sValueName, ProwlKeyExists
	If IsNull (ProwlKeyExists) Then
		WScript.Echo "Key does not exist"
	Else
		oShell.RegDelete sRegKey
		WScript.Echo "Registry key deleted"
	End If
Else
	oShell.RegWrite sRegKey, ProwlAPIKey, "REG_SZ"
	WScript.Echo "API Key: " & ProwlAPIKey & " written to registry."
End If