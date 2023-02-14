dim wsh, wshOut, sShellOut, sShellOutLine
set wsh = CreateObject("WScript.Shell")
set sShellOut=wsh.exec("..\build\readdyp.exe Q").StdOut
'sShellOut = wsh.ReadAll()
WScript.Echo (sShellOut.ReadAll)