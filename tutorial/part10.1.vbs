Option Explicit
Dim obj

Set obj = createObject("wscript.shell")
msgbox obj.CurrentDirectory & "\a.txt"
