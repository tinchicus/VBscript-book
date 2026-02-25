dim objfso
dim nombre
dim archivo
dim texto

set objfso = CreateObject("Scripting.FileSystemObject")

nombre = inputbox("Ingresa un nombre para el archivo")
set archivo = objfso.createtextfile(nombre)
texto = inputbox("Ingresa un texto para el archivo")
archivo.writeline texto
archivo.close

set archivo = objfso.opentextfile(nombre)
texto = archivo.readall
msgbox texto,64,"Contenido del archivo " & nombre
archivo.close
