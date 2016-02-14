Option Explicit
Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

'fso.CreateTextFile "C:\hello.txt" Permision Denied!

'El archivo se genera a nivel de la carpeta donde se ejecuta el script.
'Por algun motivo el path absoluto no funca
fso.CreateTextFile "resultzone\hi.txt"

' If fso.FileExists("path\to\the\file") then
' fso.DeleteFolder "Path\To\File.txt"
