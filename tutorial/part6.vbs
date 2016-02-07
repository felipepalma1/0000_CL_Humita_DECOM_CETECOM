Option Explicit
Dim name
Dim age

name = InputBox ("Cual es tu nombre", "Information:", "Name goes here", 15000,1500)
If name = "" Then
 msgbox "Your name please"

ElseIf name = "Jonny" Then
 msgbox "Use your nipples"

ElseIf name <> "Jeremy" Then
 msgbox "Intruder!"

End If

age =  InputBox("What is yout age", "Information required:")
If isNumeric(age) Then
 msgbox "Ok"
ElseIf Not IsNumeric(age) Then
 msgbox "Dude Stap"
End If
