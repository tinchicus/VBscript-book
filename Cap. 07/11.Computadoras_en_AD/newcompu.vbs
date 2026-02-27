Const ADS_UF_PASSWD_NOTREQD             = &h0020
Const ADS_UF_WORKSTATION_TRUST_ACCOUNT  = &h1000
Const titulo = "Agregando equipos v.1.0"
dim computadora, descrip, ou, ruta, dominio, vCompu

Set objRootDSE = GetObject("LDAP://rootDSE")
dominio = objRootDSE.Get("defaultNamingContext")

computadora = inputbox("Ingresa el nombre del equipo",titulo)
vCompu = FindCompu(computadora, dominio)

if (vCompu="Not Found") then
        descrip = inputbox("Descripcion del equipo",titulo)
        ruta = inputbox("Ingresa la OU de destino", titulo)
        if (ruta<>"") then
                ou = ConvertirOU(ruta)
        else
                ou = "cn=Computers"
        end if
        Set objContainer = GetObject("LDAP://" & ou & "," & dominio)
        Set objComputer = objContainer.Create("Computer", "cn=" & computadora)
        objComputer.Put "sAMAccountName", computadora & "$"
        objComputer.Put "userAccountControl", _
                        ADS_UF_PASSWORD_NOTREQD Or ADS_UF_WORKSTATION_TRUST_ACCOUNT
        objComputer.Put "Description", descrip
        objComputer.SetInfo

        msgbox "Listo el pollo" & vbCrLf & "Pelada la Gallina!"
else
        Msgbox "El equipo " & computadora & " ya existe,"
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
