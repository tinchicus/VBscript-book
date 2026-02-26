Const titulo = "Eliminar grupo de AD v. 1.0"

nombre = inputbox("Ingresa el nombre del grupo",titulo)

Set objOU = GetObject("LDAP://cn=Users, dc=laboratorio,dc=local")
objOU.Delete "group", "cn=" & nombre

msgbox "Listo el pollo" & vbCrLf & "Pelada la gallina"
