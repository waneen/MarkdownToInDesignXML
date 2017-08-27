Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /c pandoc "+arguments(0),0,false
'objShell.Run "cmd /c notepad C:/Users/edit/Documents/aa.txt",0,false
