dim texto

set objfso = CreateObject("Scripting.FileSystemObject")

for each objDrive in objfso.drives
  texto = texto & "Unidad: " & objDrive.driveletter & vbCrLf
  for each objFolder in objdrive.rootfolder.subfolders
        texto = texto & objFolder.name & " / "
        texto = texto & objFolder.path & vbCrlf
  next
next

msgbox texto,64,"Datos de tus discos"
