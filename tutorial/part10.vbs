'wscript.echo "Hello"
'a = wscript.scriptname 'Se entrega el nombre de este archivo
'b = wscript.scriptfullname 'Se entrega el nombre completo de este archivo, o sea, la ruta y su nombre

' msgbox a
' msgbox b

a = Len(wscript.scriptname)
b = Len(wscript.scriptfullname)


msgbox a
msgbox b
