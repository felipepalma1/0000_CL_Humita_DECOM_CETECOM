Option Explicit
Dim obj

Set obj = createObject("wscript.shell")
obj.run "C:\Users\%username%\Desktop"
'Si quieres abrir un archivo, tiene que estar creado de antes
