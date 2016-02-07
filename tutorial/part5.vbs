'Se crea un objeto consola, que ejecute paint.

CreateObject("wscript.shell").run "mspaint.exe"

'Esto es un comentario
'* Esto es otro comentario, Microsoft recomienda usar el arterisco para identificarlos mejor.
'No hay comentarios multilinea parece

'Se crea un objeto consola y se le etiqueta en una variable
Set obj = CreateObject("wscript.shell")

obj.run "notepad.exe" 'Esto me recuerda un poco a JS
obj.run "C:\Users\fpalmac\Documents\humita" 'Este comando permite abrir carpetas directamente, no necesitas abrir explorer
'No necesitas ejecutar Notepad para abrir un documento. Solo tiene que estar previamente creado. Debe existir un comando para crear eso si
obj.run "C:\Users\fpalmac\Documents\humita\tutorial\resultzone\hello.txt"
