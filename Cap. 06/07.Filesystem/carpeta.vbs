set objfso = CreateObject("Scripting.FileSystemObject")

objfso.drives("c").rootfolder.subfolders("otraprueba").copy objfso.drives("c").rootfolder.subfolders("prueba")

objfso.drives("c").rootfolder.subfolders("otraprueba").move objfso.drives("c").path & "prueba2"

objfso.drives("c").rootfolder.subfolders("prueba2").delete

wscript.echo "Listo el pollo" & vbCrLf & "Pelada la gallina"
