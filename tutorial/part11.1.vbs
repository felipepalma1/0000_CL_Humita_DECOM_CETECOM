Sub Greet(user)
 MsgBox "Hello," & user
End Sub

x = InputBox("Ingresa tu nombre")
Call Greet(x)
