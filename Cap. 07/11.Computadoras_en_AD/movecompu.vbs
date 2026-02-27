Const titulo = "Moviendo Equipos"
dim rutanueva, rutavieja, equipo

Set objRootDSE = GetObject("LDAP://rootDSE")
Dominio = objRootDSE.Get("defaultNamingContext")

equipo = inputbox("Ingresa el nombre del equipo", titulo)
rutavieja = FindCompu(equipo, ucase(Dominio))

if (rutavieja <> "Not Found") then
        rutanueva = inputbox("Ingresa la nueva OU", titulo)
        rutanueva = ConvertirOu(rutanueva)
        equipo = "CN=" & equipo
        Set objNewOU = GetObject("LDAP://" & rutanueva & "," & Dominio)
        Set objMoveComputer = objNewOU.MoveHere(rutavieja,equipo)
        msgbox "Listo el Pollo" & vbCrLf & "Pelada la gallina"
else
        msgbox "El equipo no fue encontrado"
end if

function ConvertirOU(ByVal camino)
        dim c, texto
        c = split(camino,"/")
        for a=ubound(c) to 1 step -1
                texto = texto & "ou=" & c(a) & ","
        next
        texto = mid(texto,1,len(texto)-1)
        ConvertirOU = texto
end function

Function FindCompu(Byval CompuName, Byval Domain)

        dim vCompu

        set cn = createobject("ADODB.Connection")
        set cmd = createobject("ADODB.Command")
        set rs = createobject("ADODB.Recordset")
        cn.open "Provider=ADsDSOObject;"

        cmd.activeconnection=cn
        cmd.commandtext="SELECT ADSPath FROM 'GC://" & Domain & _
           "' WHERE sAMAccountName = '" & CompuName & "$'"

        set rs = cmd.execute
        if err<>0 then
                 FindCompu="Error conectandose a la base del AD:" & err.description
        else
                if not rs.BOF and not rs.EOF then
                        rs.MoveFirst
                        vCompu = Replace(rs(0),"GC://","LDAP://")
                        FindCompu = vCompu
                else
                        FindCompu = "Not Found"
                end if
        end if
        cn.close
end function
