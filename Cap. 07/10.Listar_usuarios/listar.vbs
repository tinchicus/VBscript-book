Const titulo = "Listas de usuario v. 1.0"

Dim oRootLDAP, Dominio
Dim nombre, departamento, pais, descripcion, empresa, texto
Dim objFSOut, objStreamout, fileout

ahora = now
ano = datepart("yyyy",ahora)
mes = datepart("m",ahora)
dia = datepart("d",ahora)
hora = datepart("h",ahora)
minuto = datepart("n",ahora)
segundo = datepart("s",ahora)
ahora = ano & mes & dia & hora & minuto & segundo

fileout = "Listado -" & ahora & ".log"

Set objFSOut = CreateObject("Scripting.FileSystemObject")
Set objStreamout = objFSOut.CreateTextFile(fileout, 1, false)

Set oRootLDAP = GetObject("LDAP://rootDSE")
Dominio = oRootLDAP.get("defaultNamingContext")
Set oContenedor = GetObject("LDAP://" & Dominio)

objStreamout.WriteLine("Inicio: " & date & " " & time)
objStreamout.WriteLine("cuenta;nombre completo;departamento;pais;descripcion;empresa")

listUsers(oContenedor)

objStreamout.write "Final: " & date() & " " & time()
objStreamout.close()
Set objStreamout = nothing
Set objFSOut = nothing

msgbox "Listo el pollo" & vbCrLf & "Pelada la gallina"

sub listUsers(oObjeto)
dim oUser

for each oUser in oObjeto
        select case lcase(oUser.class)
                case "user"
                        cuenta=oUser.get("sAMAccountname")
                        nombre=ObtenInfo(cuenta,"DisplayName",Dominio)
                        departamento=ObtenInfo(cuenta,"department",Dominio)
                        pais=ObtenInfo(cuenta,"c",Dominio)
                        descripcion=oUser.get("description")
                        empresa=ObtenInfo(cuenta,"company",Dominio)
                        texto = cuenta & ";" & nombre & ";" & departamento & ";" & _
                                pais & ";" & descripcion & ";" & empresa
                        objStreamout.writeline texto
                case "organizationalunit", "container"
                        listUSers(GetObject("LDAP://" &oUser.get("distinguishedName")))
        end select
next
end sub

function ObtenInfo(oUser,oCampo, cDominio)
        set cn = createobject("ADODB.Connection")
        set cmd = createobject("ADODB.Command")
        set rs = createobject("ADODB.Recordset")
        cn.open "Provider=ADsDSOObject;"

        cmd.activeconnection=cn
        cmd.commandtext="SELECT " & oCampo & " FROM 'LDAP://" & cDominio & _
           "' WHERE sAMAccountName = '" & oUser & "'"
        
        set rs = cmd.execute
        if err<>0 then
                 FindUser="Error conectandose a la base del AD:" & err.description
        else
                if (not rs.BOF and not rs.EOF) AND (rs(0)<>"") then
                        ObtenInfo=rs(0)
                else
                        ObtenInfo="N/A"
                end if
        end if
        cn.close
end function
