Option Explicit

Dim filler, x

set x = createObject("wscript.shell")
filler = inputbox("What to search?")

x.run "firefox.exe"
wscript.sleep 2000
x.sendkeys filler
x.sendkeys "{enter}"
