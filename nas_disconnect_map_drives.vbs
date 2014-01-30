'NAStools v0.1
'Written by Dickson Chow
'https://github.com/dicksondickson
'January 30, 2014

'This script will disconnect the currently mapped drives E and F
Dim objNetwork
   Set objNetwork = CreateObject("WScript.Network")
objNetwork.RemoveNetworkDrive "E:"
objNetwork.RemoveNetworkDrive "F:"
