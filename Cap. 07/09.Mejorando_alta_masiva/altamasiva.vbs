Const ADS_PROPERTY_APPEND = 3
Const ADS_PORPERTY_UPDATE = 2
Const Dominio = "dc=laboratorio,dc=local"
Const udominio = "laboratorio.local"
Const titulo = "Alta de usuario v. 1.2"
Const filein = "usuarios.csv"

ahora = now
ano = datepart("yyyy",ahora)
mes = datepart("m",ahora)
dia = datepart("d",ahora)
hora = datepart("h",ahora)
minuto = datepart("n",ahora)
segundo = datepart("s",ahora)
ahora = ano & mes & dia & hora & minuto & segundo

fileout = "Creados -" & ahora & ".log"

Dim objFSOut, objStreamout, fileout,textolog
Dim objFSO, objStream, linea, vuelta

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objStream = objFSO.OpenTextFile(filein, 1, false)

Set objFSOut = CreateObject("Scripting.FileSystemObject")
Set objStreamout = objFSOut.CreateTextFile(fileout, 1, false)

objStreamout.WriteLine("Inicio: " & date() & " " & time())
do while objStream.AtEndOfStream <> true
        linea = objStream.readline
        if vuelta > 0 then CrearUsuario(linea)
        vuelta = vuelta + 1
loop

objStreamout.write "Final: " & date() & " " & time()
objStreamout.close()
Set objStreamout = nothing
Set objFSOut = nothing

objStream.close()
Set objStream = nothing
Set objFSO = nothing

msgbox "Listo el pollo" & vbCrLf & "Pelada la gallina"

sub CrearUsuario(datos)

d = split(datos,";")

cuenta = d(0)
nombre = d(1)
apellido = d(2)
uDisplay = apellido & ", " & nombre

chequeo = FindUser(cuenta,uDisplay,ucase(Dominio))

if (chequeo="Not Found") then

empresa = d(3)
descrip = d(4)
pais = d(5)
departamento = d(6)

select case lcase(departamento)
        
case "presidencia"
        g="presidencia"
        ou = ConvertirOU("laboratorio.local/Tinchicus/Usuarios/Argentina/Presidencia")
case "it"
        g="it"
        ou = ConvertirOU("laboratorio.local/Tinchicus/Usuarios/Argentina/IT")
case "comercio exterior"
        g="comex"
        ou = ConvertirOU("laboratorio.local/Tinchicus/Usuarios/Argentina/Comercio Exterior")
case "vendors"
        g="vendors"
        ou = ConvertirOU("laboratorio.local/Tinchicus/Usuarios/Argentina/Vendors")
case "ventas"
        g="ventas"
        ou = ConvertirOU("laboratorio.local/Tinchicus/Usuarios/Argentina/Ventas")
case "call center"
        g="call_center"
        ou = ConvertirOU("laboratorio.local/Tinchicus/Usuarios/Argentina/Call Center")
case else
        g="Vendors"
        ou="cn=Users"   
end select

Set oRootLDAP = GetObject("LDAP://rootDSE")
Set oContenedor = GetObject("LDAP://" & ou & "," & Dominio)
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

grupo=FindGroup(g, Dominio)
Set objGroup = GetObject("LDAP://" & grupo )
u=FindGroup(cuenta,Dominio)
objGroup.PutEx ADS_PROPERTY_APPEND, "member", Array(u)
objGroup.SetInfo

randomize timer

dim texto

do while a<12
        if (a=0) then 

        caracter = int(rnd * 24) + 65
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
oNuevoUsuario.Put "pwdLastSet", 0

oNuevoUsuario.put "userAccountControl", 544
oNuevoUsuario.SetInfo

textolog = "Se creo la cuenta: " & cuenta & vbCrLf
textolog = textolog & "El nombre completo: " & uDisplay & vbCrLf
textolog = textolog & "La contraseña es: " & uPassword & vbCrLf
textolog = textolog & "========================================" & vbCrLf

objStreamout.Write textolog

else

textolog = "Se encontro el siguiente objeto: " & chequeo &vbCrLf
textolog = textolog & "========================================" & vbCrLf

objStreamout.Write textolog

end if

end sub

Function FindUser(Byval UserName, ByVal CanonName, Byval Domain)

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

Function FindGroup(Byval GroupName, Byval Domain)

        dim vGrupo

        set cn = createobject("ADODB.Connection")
        set cmd = createobject("ADODB.Command")
        set rs = createobject("ADODB.Recordset")
        cn.open "Provider=ADsDSOObject;"

        cmd.activeconnection=cn
        cmd.commandtext="SELECT ADsPath FROM 'GC://" & Domain & _
           "' WHERE sAMAccountName = '" & GroupName & "'"

        set rs = cmd.execute
        if err<>0 then
                 FindUser="Error conectandose a la base del AD:" & err.description
        else
                rs.MoveFirst
                vGrupo = Replace(rs(0),"GC://","")
                FindGroup=vGrupo
        end if
        cn.close
end function 

function ConvertirOU(ByVal camino)
        dim c, texto
        c = split(camino,"/")
        for a=ubound(c) to 1 step -1
                texto = texto & "ou=" & c(a) & ","
        next
        texto = mid(texto,1,len(texto)-1)
        ConvertirOU = texto
end function
