Const Dominio = "dc=laboratorio,dc=local"
Const udominio = "laboratorio.local"
Const titulo = "Alta de usuario v. 1.0"

cuenta = inputbox("Ingresa la cuenta",titulo)
nombre = inputbox("Ingresa el nombre",titulo)
apellido = inputbox("Ingresa lel apellido",titulo)
uDisplay = apellido & ", " & nombre
chequeo = FindUser(cuenta,uDisplay,ucase(Dominio))

if (chequeo="Not Found") then

empresa = inputbox("Ingresa la empresa",titulo)
descrip = inputbox("Ingresa la descripcion",titulo)
pais = inputbox("Ingresa el pais",titulo)
departamento = inputbox("Ingresa el departamento",titulo)

Set oRootLDAP = GetObject("LDAP://rootDSE")
Set oContenedor = GetObject("LDAP://CN=Users," & Dominio)
Set oNuevoUsuario = oContenedor.Create("User","cn=" & chr(34) & uDisplay & chr(34))

oNuevoUsuario.put "sAMAccountName", lcase(cuenta)
oNuevoUsuario.put "givenName", nombre
oNuevoUsuario.put "sn", apellido
oNuevoUsuario.put "UserPrincipalName", lcase(cuenta) & "@" & uDominio
oNuevoUsuario.put "cn", uDisplay
oNuevoUsuario.put "DisplayName", uDisplay
oNuevoUsuario.put "company", empresa
oNuevoUsuario.put "c", pais
oNuevoUsuario.put "department", departamento
oNuevoUsuario.put "Description", descrip

oNuevoUsuario.SetInfo

randomize timer

dim texto

do while a<12
        if (a=0) then 

        caracter = int(rnd * 25) + 90
        texto = ucase(chr(caracter))
        a = a + 1

        else
        
        caracter = int(rnd * 122) + 1
        if ((caracter>47 and caracter<58) or (caracter>96 and caracter<123)) then
                texto = texto & chr(caracter)
                a = a + 1
        end if

        end if           
loop

uPassword = texto
oNuevoUsuario.SetPassword uPassword
oNuevoUsuario.put "userAccountControl", 544
oNuevoUsuario.SetInfo

msgbox "Se creo el usuario y la pass es: " & uPassword

else

msgbox "Se encontro el siguiente objeto: " & chequeo,,titulo

end if

Function FindUser(Byval UserName, ByVal CanonName, Byval Domain)


        set cn = createobject("ADODB.Connection")
        set cmd = createobject("ADODB.Command")
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
                        FindUser=UserName
                 else
                        cmd.commandtext="SELECT ADsPath FROM 'GC://" & Domain & _
                        "' WHERE cn = '" & CanonName & "'"

                        set rs = cmd.execute
                        if not rs.BOF and not rs.EOF then
                                FindUser=CanonName
                        else
                                FindUser = "Not Found"
                        end if
                end if
        end if
        cn.close
end function
