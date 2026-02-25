dim texto

set objfso = CreateObject("Scripting.FileSystemObject")

for each objDrive in objfso.drives
        texto = texto & "Unidad: " & objDrive.driveletter & vbCrLf
        texto = texto & "Sistema de archivos: " & objDrive.filesystem & vbCrLf
        texto = texto & "Serial number: " & objDrive.serialnumber & vbCrLf
        texto = texto & "Nombre del volumen: " & objDrive.volumename & vbCrLf
        texto = texto & "Tamaño: " & objDrive.totalsize & vbCrLf
        texto = texto & "Espacio libre: " & objDrive.freespace & vbCrLf
next

msgbox texto,64,"Datos de tus discos"
