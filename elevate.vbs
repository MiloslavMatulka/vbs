' Starts a process elevated, command line utility

' elevate.vbs /f:filename [/p:parameters] [/d:dir] [/v:verb] [/w:window]

' /f:filename   : Specifies the filename (pathname) to execute.
' /p:parameters : Specifies arguments for the executable.
' /d:dir        : Specifies the working directory.
' /v:verb       : Specifies the operation to execute (runas=default/open/edit/print).
' /w:window     : Specifies view mode application window (1=normal, 0=hide, 2=Min, 3=max, 4=restore, 5=current, 7=min/inactive, 10=default).

Set objArgs = WScript.Arguments.Named

strFile = objArgs.Item("f")
strParams = objArgs.Item("p")
strDir = objArgs.Item("d")
If objArgs.Item("v") = Empty Then
    strVerb = "runas"
Else
    strVerb = objArgs.Item("v")
End If
intWindow = CInt(objArgs.Item("w"))

Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute strFile, strParams, strDir, strVerb, intWindow
