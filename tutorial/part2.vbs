
Dim a

a=MsgBox("Pandora Radio",vbExclamation+vbOkCancel+vbDefaultButto­n2,"Error:")

If a = vbOK then msgbox "We'll try an fix the problem." ,vbInformation
ElseIf a = vbCancel then msgbox "Closing the program." ,vbCritical
End If
