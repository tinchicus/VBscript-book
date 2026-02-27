 Const titulo = "Eliminando Equipos"
dim equipo, ruta, dominio, aviso

Set objRootDSE = GetObject("LDAP://rootDSE")

equipo = inputbox("Ingresa el equipo", titulo)
dominio = objRootDSE.Get("defaultNamingContext")
ruta = FindCompu(equipo, Dominio)

if (ruta<>"Not Found") then
        aviso = msgbox("Esta Seguro?",1,titulo)
        if (aviso = 1) then
                set objComputer = GetObject(ruta)
                objComputer.DeleteObject(0)
                msgbox equipo & " Was Destroyed!!!",,titulo     
        end if
else
        msgbox "No se encontro el equipo",,titulo
end if

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
