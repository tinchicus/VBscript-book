Const ADS_PROPERTY_DELETE = 4 
Const titulo = "Quitar usuarios a un grupo v. 1.0"
Const uDominio = "dc=laboratorio, dc=local"

do while a < 12

grupo = inputbox("Ingresa el nombre del grupo o q para salir", titulo)

if lcase(grupo) = "q" then exit do

Set objGroup = GetObject("LDAP://cn=" & chr(34) & grupo & chr(34) &  ",cn=Users," _
 		& uDominio) 
cuenta = inputbox("Ingresa la cuenta a quitar", titulo)

chequeo = FindUser(cuenta, uDominio)

if (chequeo <> "Not Found") then 
	objGroup.PutEx ADS_PROPERTY_DELETE, "member", Array(chequeo)
	objGroup.SetInfo
else
	msgbox "El usuario no fue encontrado",64,titulo
end if

loop

msgbox "Colorin, Colorado" & vbCrLF & "Este script ha acabado"

Function FindUser(Byval UserName, Byval Domain)

	dim vUsuario

	set cn = createobject("ADODB.Connection")
	set cmd = createobject("ADODB.Command")
	set rs = createobject("ADODB.Recordset")
	cn.open "Provider=ADsDSOObject;"

	cmd.activeconnection=cn
	cmd.commandtext="SELECT ADsPath FROM 'GC://" & Domain & _
	   "' WHERE sAMAccountName = '" & UserName & "'"

	set rs = cmd.execute
	if err<>0 then
		 FindUser="Error conectandose a la base del AD:" & err.description
	else

		 if not rs.BOF and not rs.EOF then
			rs.MoveFirst
			vUsuario = Replace(rs(0),"GC://","")
			FindUser=vUsuario
		 else
			FindUser = "Not Found"
	 	end if
	end if
	cn.close
end function 
