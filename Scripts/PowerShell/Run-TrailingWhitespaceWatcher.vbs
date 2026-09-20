Option Explicit

Dim fileSystem
Dim shell
Dim scriptPath
Dim command

Set fileSystem = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptPath = fileSystem.BuildPath(fileSystem.GetParentFolderName(WScript.ScriptFullName), "TrailingWhitespaceWatcher.ps1")
command = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -STA -File """ & scriptPath & """"
shell.Run command, 0, False