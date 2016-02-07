Option Explicit
Dim a
a=MsgBox("Guess a button", vbAbortRetryIgnore)
if a=vbRetry Or a=vbAbort then
 msgbox "correct"
Else
 msgbox "wrong"
End if
