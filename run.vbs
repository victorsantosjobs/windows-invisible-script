Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c foto.bat"
oShell.Run strArgs, 0, false