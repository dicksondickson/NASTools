'NAStools v0.1
'Written by Dickson Chow
'https://github.com/dicksondickson
'January 30, 2014

'Requires IPMIUtil
'http://ipmiutil.sourceforge.net/

'This script will disconnect the mapped drives and shutdown the NAS via IPMI.
'Go through the script and change hostname/ipaddress, user name, password etc where specified.

'Disconnect map Drives
Dim objNetwork
Set objNetwork = CreateObject("WScript.Network")
objNetwork.RemoveNetworkDrive "E:"
objNetwork.RemoveNetworkDrive "F:"

'Shutdown the NAS via IPMI
Sub pwroff( cmd )
    Dim wshShell
    Set wshShell = CreateObject( "WScript.Shell" )
	'Change the path so the script can locate ipmiutil.exe
    wshShell.Run "c:\path_to_ipmi_install\ipmiutil.exe " & ( cmd ), 0, False
    Set wshShell = Nothing
End Sub

'Change the hostname/ipaddress, user and password
pwroff "reset -D -N HOSTNAME -U USER -P PASSWORD"

'Show dialogue box to show the script has executed.
x=msgbox("NAS going offline" ,0, "pr0n NAS")