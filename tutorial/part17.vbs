Option Explicit
Dim obj, desk

Set obj = createObject("wscript.shell")
desk = obj.specialFolders("Desktop")

obj.run desk
