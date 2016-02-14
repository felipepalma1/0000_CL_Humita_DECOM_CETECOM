'SubRutinas!!
'Llamando desde el principio HOISTING!
Call Saludo

' Inicio subrutina
Sub Saludo

'Hello world
msgbox "Hello!"
' Final de subrutina
End Sub

'No pasara nada hasta que se invoque

Call Saludo 'Invocando subrutina


Sub Nice(x) 'Puedes ingresar argumentos!
 msgbox "Hello", 20, x
End Sub


Sub Nic(x, word) 'Puedes ingresar argumentos!
 msgbox "Hello", x, word
End Sub

Call Nic("Welcome nice person", "Nice")
