' Start HRIS Downloads Watcher.vbs
' -----------------------------------------------------------------------------
' Launches watch_downloads.py hidden (no console window) via pythonw.exe.
' Deployed to:  %LOCALAPPDATA%\hris-downloads-watcher\
' A copy of this file also lives in the Startup folder so it runs at logon:
'   %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
' The watcher itself enforces single-instance, so running it twice is harmless.
'
' To stop it: set "enabled": false in watch_downloads.config.json (exits within
' one poll), or delete the Startup copy of this .vbs and taskkill the pythonw
' process. See hris-dashboard RESUME.md (27 Aug 2026 session).
' -----------------------------------------------------------------------------
Option Explicit
Dim sh, fso, deployDir, target, pyw

Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' The deployed watcher always lives here, regardless of where this .vbs runs from.
deployDir = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\hris-downloads-watcher"
target    = deployDir & "\watch_downloads.py"

If Not fso.FileExists(target) Then
    ' Fall back to the folder this .vbs is sitting in (e.g. first-run from the repo copy).
    deployDir = fso.GetParentFolderName(WScript.ScriptFullName)
    target    = deployDir & "\watch_downloads.py"
End If

If Not fso.FileExists(target) Then
    WScript.Quit 1
End If

pyw = "pythonw.exe"
sh.CurrentDirectory = deployDir
' 0 = hidden window, False = don't wait
sh.Run """" & pyw & """ """ & target & """", 0, False
