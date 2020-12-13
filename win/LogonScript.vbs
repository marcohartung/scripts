'LogonScript.vbs
' do every thing that must be done after logon
' Marco Hartung, 08.12.2020
' Version 1.0.0

'This work is licensed under the Creative Commons Attribution 4.0 International License.
'To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/ or send a letter to Creative Commons, PO Box 1866, Mountain View, CA 94042, USA.

Option Explicit

Dim localHost, iNetHost
Dim iNetOk, localNetOk
Dim i
Dim ToolsLaunch_PreNetwork(10,2)
Dim ToolsLaunch_PreVPN(10,2)
Dim ToolsLaunch_PostVPN(10,2)
Dim NetworkShares(10,2)

' -------------------------- ------------------------------------------------------------
'                                    USER CONFIG
' -------------------------- ------------------------------------------------------------

' ip or host in local network 
localHost = "192.168.2.1"

' internet ip or host
iNetHost = "google.de"

'ToolsLaunch_xxxx(x,0)  executable
'ToolsLaunch_xxxx(x,1)  args
'ToolsLaunch_xxxx(x,2)  0 = hide window; 1 = show window

ToolsLaunch_PreNetwork(0,0) = "notepad.exe"
ToolsLaunch_PreNetwork(0,1) = ""
ToolsLaunch_PreNetwork(0,2) = "1"

'ToolsLaunch_PreVPN(0,0) = "Path to exe"
'ToolsLaunch_PreVPN(0,1) = "ARgs"

'ToolsLaunch_PostVPN(0,0) = "Path to exe"
'ToolsLaunch_PostVPN(0,1) = "ARgs"

'NetworkShares(0,0) = "T:"
'NetworkShares(0,1) = "\\192.168.2.1\user"

' -------------------------- ------------------------------------------------------------
'                               END USER CONFIG
' -------------------------- ------------------------------------------------------------

' -------------------------- CODE  --------------------------

' Launch tools ( no network)
LaunchTools ToolsLaunch_PreNetwork 

' Wait for Network
For i = 0 to 20
	iNetOk = Ping( iNetHost )
	localNetOk = Ping( localHost  )
  
	If iNetOk = true Then
		Exit For
	End If
Next	
'wscript.echo "iNeT: " & iNetOK & " - localNetOk: " & localNetOk

' Exit here if we have not network
IF iNetOk = false Then
	wscript.echo "No connection to: " & iNetHost & vbNewLine & "Propaply no network connection!" & vbNewLine &  "Quit LogonScript"
	wscript.Quit -1
End If

' Launch tools (internet OK;  no private network)
LaunchTools ToolsLaunch_PreVPN 

' Should we launch VPN
IF (iNetOk = true) AND (localNetOk = false) Then

	Call ShowVPN

	For i = 0 to 20
		localNetOk = Ping( localHost )
		If localNetOk = true Then
			Exit For
		End If
	Next

	IF localNetOk = false Then
		wscript.echo "No connection to: " & localHost & vbNewLine & "Propaply no VPN connection!" & vbNewLine &  "Quit LogonScript"
		wscript.Quit -1
	End If
End If

' add network shares
AddNetworkShares NetworkShares

' Launch tools (full network)
LaunchTools ToolsLaunch_PreVPN 

wscript.Quit 0

' -------------------------- FUNCTIONS  --------------------------

Function Ping( host )
	Dim result
	Dim shell, shellexec

	Set shell = WScript.CreateObject("WScript.Shell")
	Set shellexec = shell.Exec("ping -n 1 -w 2000 -4 " & host) 
	result = shellexec.StdOut.ReadAll
	'wscript.echo result

	If InStr(result , "TTL") Then
	  Ping = true 
	Else
	  Ping = false
	End If

End Function

Sub ShowVPN()
	Set wsShell = CreateObject("WScript.Shell")
	wsShell.Run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"""
	wscript.Sleep 300
	wsShell.AppActivate "Cisco AnyConnect Secure Mobility Client"
End Sub

Sub LaunchTools( strTools )

	For i = 0 To ubound( strTools, 1)
		If Len(strTools(i,0)) > 0 Then
			Dim wsShell, strCmd, strArgs
			Set wsShell = CreateObject("WScript.Shell")
			strCmd = chr(34) & wsShell.ExpandEnvironmentStrings(strTools(i,0)) & chr(34)
			If Len(strTools(i,1)) > 0 Then
				strArgs = chr(32) & chr(34) &  strTools(i,1) & chr(34)
			End If
			'wscript.echo strCmd
			'wscript.echo strArgs
			wsShell.Run strCmd & strArgs, strTools(i,2)
			Set wsShell = Nothing
		End If
	Next

End Sub

Sub AddNetworkShares( strShares )
' connect to or fix broken NetShares
' source: https://ss64.com/vb/syntax-mapdrive.html
	For i = 0 To ubound( strShares, 1)
		If Len(strShares(i,0)) > 0 Then
			Dim objNetwork, objDrives, objReg, n
			Dim strLocalDrive, strRemoteShare, strShareConnected
			Dim bolFoundExisting, bolFoundRemembered
			Const HKCU = &H80000001
			
			strLocalDrive = strShares(i,0)
			strRemoteShare = strShares(i,1)
			'wscript.echo " - Mapping: " + strLocalDrive + " to " + strRemoteShare
			
			Set objNetwork = CreateObject("WScript.Network")
			' Loop through the network drive connections and disconnect any that match strLocalDrive
			Set objDrives = objNetwork.EnumNetworkDrives
			If objDrives.Count > 0 Then
			  For n = 0 To objDrives.Count-1 Step 2
			    If objDrives.Item(n) = strLocalDrive Then
			      strShareConnected = objDrives.Item(n+1)
			      objNetwork.RemoveNetworkDrive strLocalDrive, True, True
			      n=objDrives.Count-1
			      bolFoundExisting = True
			    End If
			  Next
			End If

			' If there's a remembered location (persistent mapping) delete the associated HKCU registry key
			If bolFoundExisting <> True Then
			  Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
			  objReg.GetStringValue HKCU, "Network\" & Left(strLocalDrive, 1), "RemotePath", strShareConnected
			  If strShareConnected <> "" Then
			    objReg.DeleteKey HKCU, "Network\" & Left(strLocalDrive, 1)
			    bolFoundRemembered = True
			  End If
			End If

			'Now actually do the drive map (not persistent)
			Err.Clear
			On Error Resume Next
			objNetwork.MapNetworkDrive strLocalDrive, strRemoteShare, False

			'Error traps
			If Err <> 0 Then
			  Select Case Err.Number
			    Case -2147023694
			      'Persistent connection so try a second time
			      On Error Goto 0
			      objNetwork.RemoveNetworkDrive strLocalDrive, True, True
			      objNetwork.MapNetworkDrive strLocalDrive, strRemoteShare, False
			      WScript.Echo "Second attempt to map drive " & strLocalDrive & " to " & strRemoteShare
			    Case Else
			      On Error GoTo 0
			      WScript.Echo " - ERROR: Failed to map drive " & strLocalDrive & " to " & strRemoteShare
			  End Select
			  Err.Clear
			End If

		End If
	Next
End Sub
