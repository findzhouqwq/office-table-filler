Option Explicit
' 静默启动图形安装程序，不显示命令框。
Dim fso, shell, baseDir, scriptPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fso.BuildPath(baseDir, "TableFillerSetup.ps1")

If Not fso.FileExists(scriptPath) Then
    MsgBox "找不到安装脚本 TableFillerSetup.ps1，请确认文件完整。", 16, "安装程序"
    WScript.Quit 1
End If

shell.Run "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """", 0, False
