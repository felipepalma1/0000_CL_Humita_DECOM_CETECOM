Sub Refresh
Set objREAD = fso.OpenTextFile(CD,vbRead)
current=objREAD.ReadAll
objREAD.Close
End Sub

Sub Main
Do
create=InputBox(current ,"Text Editor: [*] to quit")
If create = "*" then
wscript.quit
ElseIf create = "" then
Set objWRITE = fso.OpenTextFile(CD,vbAppend,1)
objWRITE.writeblanklines(1)
objWRITE.Close
Call Refresh
Else
Set objWRITE = fso.OpenTextFile(CD,vbAppend,1)
objWRITE.write create
objWRITE.Close
Call Refresh
End If
Loop
End Sub

'---------------------------------
'WRITTING PROCESS
'---------------------------------

If Not fso.FileExists(CD) then
Call Main
Else
If fso.GetFile(CD).Size = 0 Then
Call Main
Else
Call Refresh
Call Main
End If
End If
Wscript.Quit
________________________________________­______________

Know the Basics:
----------------------------------------­----------------------------------------­----------
Set objFso = CreateObject("Scripting.FileSystemObject­")
Set x = objFso.OpenTextFile("FILE-LOCATION", 8)

x.Write "" = add directly onto the end of your file.
x.WriteLine "" = add directly onto the end of your file, followed by a line.
x.WriteBlankLines(#) = add blank lines to the end of your file.
