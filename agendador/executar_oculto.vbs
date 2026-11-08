' Roda o relatorio sem abrir janela de console: o Agendador de Tarefas chama este arquivo com o wscript.exe.
Set sh = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
sh.CurrentDirectory = fs.GetParentFolderName(fs.GetParentFolderName(WScript.ScriptFullName))
codigo = sh.Run("python -m relatorio", 0, True)
WScript.Quit codigo
