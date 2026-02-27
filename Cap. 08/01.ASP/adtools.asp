<%@ language="vbscript" %>

<%
user=request.form("usuario")

Set objRootDSE = GetObject("LDAP://rootDSE")
dominio = objRootDSE.Get("defaultNamingContext")

vuser = FindUser(user, dominio)

if (vuser = "Not Found") then

texto =  user & " no encontrado en el AD."

else

Set objUser = GetObject("LDAP://" & vuser)
objUser.IsAccountLocked = False
objUser.SetInfo

texto = user & " fue desbloqueado"

end if

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
	
%>
<html>
<head><title>AD Tools</title></head>

<body>
<form id="form1" name="form1" method="POST" action="">
	Desbloqueo de usuarios<br>
	Ingresa el usuario: <input id="usuario" name="usuario" value=""><br>
	<button type="submit">Desbloquea el usuario</button>
</form>
<% response.write(texto) %>
</body>
</html>
