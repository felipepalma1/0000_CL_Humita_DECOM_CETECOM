Option Explicit

Dim ie, ipf

Set ie = CreateObject("InternetExplorer.Application")
On Error Resume Next

Sub WaitForLoad
 Do While IE.Busy
  WScript.Sleep 500
 Loop
End Sub

ie.Left = 0
ie.Top = 0
ie.Toolbar = 0
ie.StatusBar = 0
ie.Height = 120
ie.Width = 1020
ie.Resizable = 0
