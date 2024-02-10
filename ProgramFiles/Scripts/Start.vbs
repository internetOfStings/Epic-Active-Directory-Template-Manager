Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File ""Epic-ADTemplateManager.ps1""", 0, True
Set objShell = Nothing

