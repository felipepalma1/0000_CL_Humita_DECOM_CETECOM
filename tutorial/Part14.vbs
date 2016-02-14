Option Explicit
Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

fso.CopyFile "C:\Users\%username%\Desktop\a.txt", "C:\Users\%username%\Desktop\b.txt"
'fso.CopyFolder
'fso.MoveFile
'fso.MoveFolder
