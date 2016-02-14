Option Explicit
Dim ie
'Esto funciona en Win 10. IE esta en sistema

Set ie = CreateObject("InternetExplorer.Application")
ie.Navigate "http://www.facebook.com"
ie.Visible = 1
ie.Toolbar = 1
ie.StatusBar = 1
 
