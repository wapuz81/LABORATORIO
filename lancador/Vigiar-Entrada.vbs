' Roda o Vigiar-Entrada.ps1 sem abrir janela (chamado pela tarefa agendada)
Dim pasta
pasta = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
CreateObject("WScript.Shell").Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & pasta & "\Vigiar-Entrada.ps1""", 0, False
