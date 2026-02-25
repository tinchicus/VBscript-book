dim objfso
dim archivo
dim texto

set objfso = CreateObject("Scripting.FileSystemObject")

for each archivo in objFso.drives("c").rootfolder.subfolders("prueba").files
        texto = texto & archivo.name & vbTab
        texto = texto & archivo.size & vbTab
        texto = texto & archivo.type & vbCrLf
next

wscript.echo texto
