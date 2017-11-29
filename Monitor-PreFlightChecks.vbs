Option Explicit

Dim oRegistry, KeyPath, ValueName, ProwlAPIKey, ProwlNotifications, ProwlDisable

Const HKCU=&H80000001

Set oRegistry=GetObject("winmgmts:\\.\root\default:StdRegProv")

KeyPath="Software\boredazfcuk\mining"
ValueName="ProwlAPIKey"
oRegistry.GetStringValue HKCU, KeyPath, ValueName, ProwlAPIKey, ProwlDisable
	
If Not IsNull(ProwlAPIKey) Then
	ProwlNotifications=True
End If

'----- Change line below to True to disable Prowl notifications for this script only -----
ProwlDisable=False

'----- If Prowl Notifications are enabled -----
If ((ProwlNotifications) And (Not ProwlDisable))Then
	SendProwlNotification "2","Monitor-PreFlightChecks","Mining Rig Started"
End If

Sub SendProwlNotification(Priority, Application, Description)
	Dim oHTTP
	Set oHTTP=CreateObject("Microsoft.XMLHTTP")  
	oHTTP.Open "Get", "https://prowl.weks.net/publicapi/add?" & "apikey=" & ProwlAPIKey & "&priority=" & Priority & "&application=" & Application & "&event=" & Date() & " " & Time()  & "&description=" & Description ,false  
	oHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"  
	oHTTP.Send  
End Sub