Const uDominio = "dc=laboratorio, dc=local"
Const titulo = "Eliminacion de usuarios v. 1.0"

cuenta = inputbox("Ingresa el usuario", titulo)

chequeo = FindUser(cuenta, uDominio)

if chequeo <> "Not Found" then

d = split(chequeo, ",")
c = d(0) & "," & d(1)
cuenta = replace(c,"\","")
cuenta = mid(cuenta,1,3) & chr(34) & mid(cuenta,4,len(cuenta)) & chr(34)
for a=2 to ubound(d)
        ou = ou & "," & d(a)
next
ou = lcase(mid(ou,2,len(ou)))

Set objOU = GetObject("LDAP://" & ou)
objOU.Delete "user", cuenta

msgbox "Usuario terminado."

else

msgbox "El usuario no fue encontrado."

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
