Const ADS_UF_ACCOUNTENABLE = 544
Const uDominio = "dc=laboratorio, dc=local"
Const titulo = "Desbloqueo de usuarios v. 1.0"

cuenta = inputbox("Ingresa el usuario", titulo)

chequeo = FindUser("b-" & cuenta, uDominio)

if chequeo <> "Not found" then
        
        Set objUser = GetObject(chequeo)
        
        objUser.Put "userAccountControl", ADS_UF_ACCOUNTENABLE
        objUser.SetInfo
        
        objUser.AccountExpirationDate = cdate("1/1/1970")
        objUser.SetInfo
        objUser.put "sAMAccountName", lcase(cuenta)
        objUser.put "UserPrincipalName", lcase(cuenta & "@laboratorio.local")
        objUser.put "Description", "usuario reactivado"
        objUser.SetInfo
else
        msgbox "Usuario no encontrado"
end if

msgbox "Script finalizado"

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
                        vUsuario = Replace(rs(0),"GC://","LDAP://")
                        FindUser=vUsuario
                 else
                        FindUser = "Not Found"
                end if
        end if
        cn.close
end function
