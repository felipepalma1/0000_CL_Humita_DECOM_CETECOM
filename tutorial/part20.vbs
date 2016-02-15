Option Explicit
Dim fso, oFile, route, a

route = InputBox("Ingresa la ruta")

Set fso = CreateObject("Scripting.FileSystemObject")
Set oFile = fso.OpenTextFile(route, 1)

'oFile.Read(20) 'integer, numero de caracteres'
'oFile.ReadLine
'oFile.ReadAll
'oFile.Close

'MsgBox(oFile.Read(100)) 'Si el argumento es mayor a la cantidad de lineas, expresara la linea entera'
                        'Si el argumento es mayor a la cantidad de caracteres de una linea, seguira de largo

                        'Usa un loop para leer un archivo una palabra a la vez'
'oFile.Close


for a = 1 to 1'00
  'MsgBox  oFile.Read(1)
  'MsgBox  oFile.ReadLine
  MsgBox oFile.ReadAll
Next
