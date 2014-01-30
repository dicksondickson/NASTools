'NAStools v0.1
'Written by Dickson Chow
'https://github.com/dicksondickson
'January 30, 2014

'Requires IPMIUtil
'http://ipmiutil.sourceforge.net/

'This script will turn on the NAS via IPMI, ping it to see if it is awake and then map the drives.
'Go through the script and change hostname/ipaddress, user name, password etc where specified.

'Turn on the NAS via IPMI
Sub pwron( cmd )
    Dim wshShell
    Set wshShell = CreateObject( "WScript.Shell" )
	'Change the path so the script can locate ipmiutil.exe
    wshShell.Run "c:\path_to_ipmi_install\ipmiutil.exe " & ( cmd ), 0, False
    Set wshShell = Nothing
End Sub

'Change the hostname, user and password
pwron "reset -u -N HOSTNAME -U USER -P PASSWORD"

'Ping the NAS to see if it is up.
awake = False
Do Until awake = True
	'Default wait time is 2800ms. Change to your liking.
	WScript.Sleep 2800
	'Insert the IP address to your NAS below.
    Ping("IPADDRESS")
Loop

'Map Samba shares to drives.
Dim objNet, strUserName, Return
Set objNet = CreateObject("WScript.Network")
strUserName = objNet.UserName

'Change the desired drive letter and insert the hostname of your NAS and Samba shares.
Return = fnMapNetworkDrive ("E:" , "\\NAS\pr0n1")
Return = fnMapNetworkDrive ("F:" , "\\NAS\pr0n2")

'Message box to signal the user that the NAS is on and drives are mapped
x=msgbox("READY FOR FUN TIME!!!" ,0, "pr0n NAS")


'You don't need to modify anthing beyond this point.

'Map network drive function
Function fnMapNetworkDrive (Drive, Path)
Dim i, oDrives
	Set oDrives = objNet.EnumNetworkDrives
	For i = 0 to oDrives.Count - 1 Step 2 ' Find out if an existing network drive exists
	If oDrives.Item(i) = Drive Then
	'WScript.Echo "Removing drive: " & Drive
	objNet.RemoveNetworkDrive Drive, true, true
End If
	Next
	' WScript.Echo "Mapping drive: " & Drive & " to path: " & Path
	objNet.MapNetworkDrive Drive, Path
	Set i = Nothing
	Set oDrives = Nothing
	Set Drive = Nothing
	Set Path = Nothing
End Function

'Ping host function
Function Ping(address)
    Ping = False
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set objPing = objWMI.Get("Win32_PingStatus.Address='" & address & "'")
    If objPing.StatusCode = 0 Then
        Ping = True
	awake = True
    End If
End Function

