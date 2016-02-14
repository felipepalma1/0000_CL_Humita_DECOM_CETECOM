Option Explicit
Dim ie, ipf

Set ie = CreateObject("InternetExplorer.Application")

Sub WaitForLoad
  Do While IE.Busy
   WScript.Sleep 500
  Loop
End Sub

'ie.Left = 0
