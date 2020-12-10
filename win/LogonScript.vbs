'LogonScript.vbs
' do every thing that must be done after logon
' Marco Hartung, 08.12.2020

' "WshShell.Run strScript, 0" 0 = hide window, 1 = show window (useful for debugging)
Option Explicit

Dim localHost
Dim iNetHost
Dim iNetOk
Dim localNetOk
Dim i

localHost = "192.168.2.5"
iNetHost = "google.de"

' Wait for Network
For i = 0 to 2
	iNetOk = Ping( iNetHost )
	localNetOk = Ping( localHost  )
  
	If iNetOk = true Then
		Exit For
	End If
Next
	
wscript.echo "localNetOk: " & localNetOk
wscript.echo "iNeT: " & iNetOK

' Should we launch VPN
IF (iNetOk = true) AND (localNetOk = false) Then

	ReConnectVPN WshShell, strPassword

	For i = 0 to 20
		iNetOk = Ping( localHost )
		If iNetOk = true Then
			Exit For
		End If
	Next

	IF localNet = false Then
		wscript.echo "No connection to local net: " & localHost
	End If
End If

' fix broken NetShares
'Dim objNetwork
'Set objNetwork = CreateObject("WScript.Network")
'objNetwork.MapNetworkDrive "Z:", "\\server\path", false, "user", "password"

' run tools

wscript.echo "Finished!"

' -------------------------- functions  --------------------------

Function Ping( host )
	Dim result
	Dim shell, shellexec

	Set shell = WScript.CreateObject("WScript.Shell")
	Set shellexec = shell.Exec("ping -n 1 -w 2000 -4 " & host) 
	result = shellexec.StdOut.ReadAll

	wscript.echo result

	If InStr(result , "TTL") Then
	  Ping = true 
	Else
	  Ping = false
	End If

End Function

Sub ReConnectVPN(WshShell, strPassword)
	WshShell.Run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"""

	WScript.Sleep 1000

	WshShell.AppActivate "Cisco AnyConnect Secure Mobility Client"

	WshShell.SendKeys "{TAB}"
	WshShell.SendKeys "{TAB}"
	WshShell.SendKeys "{ENTER}"

	WScript.Sleep 4000

	WshShell.SendKeys strPassword
	WshShell.SendKeys "{TAB}"
	WshShell.SendKeys "{ENTER}"

	WScript.Sleep 4000

	WshShell.SendKeys "{ENTER}"
End Sub

