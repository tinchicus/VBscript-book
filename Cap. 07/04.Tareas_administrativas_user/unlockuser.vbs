Const uDominio = "dc=laboratorio, dc=local"
Const titulo = "Desbloqueo de usuario v. 1.0"

cuenta = inputbox("Ingresa la cuenta a desbloquear",titulo)

chequeo = FindUser(cuenta, uDominio)

if (chequeo = "Not Found") then

msgbox "Usuario no encontrado en el AD.",,titulo

else

Set objUser = GetObject("LDAP://" & chequeo)
objUser.IsAccountLocked = False
objUser.SetInfo

msgbox "Script finalizado"

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
