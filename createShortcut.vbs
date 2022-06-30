' Creates Windows shell links (shortcuts), command line utility.

' createShortcut.vbs /f:filename [/t:target] [/a:arguments] [/w:workingDir]
'                    [/s:style] [/i:icon,index] [/h:hotkey] [/d:description]

' /f:filename    : Specifies the .lnk file.
' /t:target      : Defines the target path and file name the shortcut points to.
' /a:arguments   : Defines the command-line parameters to pass to the target.
' /w:working dir : Defines the working directory the target starts with.
' /s:style       : Defines the window state (1=Normal, 3=Maximized, 7=Minimized).
' /i:icon,indes  : Defines the icon and optional index (file.exe or file.exe,0).
' /h:hotkey      : Defines the hotkey, a numeric value of the keyboard shortcut.
' /d:description : Defines the description (or comment) for the shortcut.

' Notes:
' - Any argument that contains spaces must be enclosed in "double quotes".
' - To prevent an environment variable from being expanded until the shortcut
'   is launched, use 2 double quotes escape character like this: ""%WINDIR""%
'   in the command-line and percent escape character like this: %%WINDIR%%
'   in a batch file

Set objArgs = WScript.Arguments.Named

Set objShell = WScript.CreateObject("WScript.Shell")
Set objLink = objShell.CreateShortcut(objArgs.Item("f"))

objLink.TargetPath = objArgs.Item("t")
objLink.Arguments = objArgs.Item("a")
objLink.WorkingDirectory = objArgs.Item("w")
objLink.WindowStyle = objArgs.Item("s")
If objArgs.Item("i") <> Empty Then
    objLink.IconLocation = objArgs.Item("i")
End If
objLink.HotKey = objArgs.Item("h")
objLink.Description = objArgs.Item("d")
objLink.Save
