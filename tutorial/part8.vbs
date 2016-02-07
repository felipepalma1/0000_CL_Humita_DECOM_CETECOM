'Enviar codigos de teclas

' Abrir notepad
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "notepad.exe", 9

' Entregar a notepad tiempo para cargar
WScript.Sleep 500

' Escribir "Hola mundo!"
WshShell.SendKeys "Hola mundo!"
WshShell.SendKeys "{ENTER}"

' Agregar la fecha
WshShell.SendKeys "{F5}"
