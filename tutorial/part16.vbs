' Create Shortcuts '
' Intente usar %username% en el path del shortcut, sin resultado'
Option Explicit
Dim obj
Dim nLink

Set obj = CreateObject("wscript.shell")

Set nLink = obj.CreateShortCut("C:\Users\fpalmac\Desktop\myNewShortcut.lnk")

nLink.TargetPath = "C:\"
nLink.Save
