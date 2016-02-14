Option Explicit
Dim ie, ipf

Set ie = CreateObject("InternetExplorer.Application")

Sub WaitForLoad
  Do While IE.Busy
   WScript.Sleep 500
  Loop
End Sub

ie.Width = 1020
ie.Resizable = 0

ie.Navigate "https://facebook.com/"

Call WaitForLoad

ie.Visible = True

ie.Document.All.Item("email").Value = "IngresarEmail"
ie.Document.All.Item("pass").Value = "IngresarPassword"
ie.Document.All.Item("u_0_n").Click
ie.FullScreen=True
