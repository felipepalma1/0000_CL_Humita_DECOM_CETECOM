Option Explicit
Dim auto
Set auto = CreateObject("wscript.shell")
auto.run "notepad.exe"
wscript.sleep 2000
auto.sendkeys "Hello there,"
auto.sendkeys "{enter}"
auto.sendkeys "hoy are you today?"
auto.sendkeys "%fs"
wscript.sleep 500
auto.sendkeys "test.txt"
auto.sendkeys "{enter}"
