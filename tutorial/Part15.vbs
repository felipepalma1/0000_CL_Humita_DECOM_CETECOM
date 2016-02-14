Option Explicit

Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

fso.CopyFile "Location", "NewLocation"
fso.CopyFolder "Location", "NewLocation"

fso.MoveFile "Location", "New Location"
fso.MoveFolder "Location", "New Location"
