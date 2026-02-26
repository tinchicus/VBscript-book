Const ADS_UF_ACCOUNTDISABLE = 2
Const uDominio = "dc=laboratorio, dc=local"
Const titulo = "Bloqueo de usuarios v. 1.0"
Const filein = "usuarios.csv"

ahora = now
ano = datepart("yyyy",ahora)
mes = datepart("m",ahora)
dia = datepart("d",ahora)
hora = datepart("h",ahora)
minuto = datepart("n",ahora)
segundo = datepart("s",ahora)
ahora = ano & mes & dia & hora & minuto & segundo

fileout = "Bloqueados -" & ahora & ".log"

Dim objFSOut, objStreamout, fileout,textolog
Dim objFSO, objStream, linea, vuelta

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objStream = objFSO.OpenTextFile(filein, 1, false)

Set objFSOut = CreateObject("Scripting.FileSystemObject")

Set objStreamout = objFSOut.CreateTextFile(fileout, 1, false)

objStreamout.WriteLine("Inicio: " & date() & " " & time())
do while objStream.AtEndOfStream <> true
        linea = objStream.readline
        if vuelta > 0 then BloquearUsuario(linea)
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

sub BloquearUsuario(usuario)

d = split(usuario,";")

cuenta = d(0)
motivo = d(1)

chequeo = FindUser(cuenta, uDominio)

if chequeo <> "Not found" then
        ahora = monthname(mes) & ", " & dia & " de " & ano
        fecha = (dateserial(ano, mes, dia)) + 1
        Set objUser = GetObject(chequeo)
        
        objUser.Put "userAccountControl", intUAC OR ADS_UF_ACCOUNTDISABLE
        objUser.SetInfo
        
        objUser.AccountExpirationDate = fecha
        objUser.SetInfo

        cuenta = "B-" & cuenta
        objUser.put "sAMAccountName", lcase(cuenta)
        objUser.put "UserPrincipalName", lcase(cuenta & "@laboratorio.local")
        objUser.put "Description", motivo & " el " & ahora
        objUser.SetInfo

        textolog = "Se bloqueo el usuario: " & cuenta &vbCrLf
        textolog = textolog & "==============================" & vbCrLf
else
        textolog = "Usuario no encontrado: " & cuenta & vbCrLf
        textolog = textolog & "==============================" & vbCrLf
end if

objStreamout.Write textolog

end sub

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
