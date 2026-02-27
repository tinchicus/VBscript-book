const uDominio = "dc=laboratorio, dc=local"
Const titulo = "Eliminacion de usuarios v. 1.0"
Const filein = "usuarios.csv"

ahora = now
ano = datepart("yyyy",ahora)
mes = datepart("m",ahora)
dia = datepart("d",ahora)
hora = datepart("h",ahora)
minuto = datepart("n",ahora)
segundo = datepart("s",ahora)
ahora = ano & mes & dia & hora & minuto & segundo

fileout = "Eliminados -" & ahora & ".log"

Dim objFSOut, objStreamout, fileout,textolog
Dim objFSO, objStream, linea, vuelta
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objStream = objFSO.OpenTextFile(filein, 1, false)

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objStream = objFSO.OpenTextFile(filein, 1, false)

Set objFSOut = CreateObject("Scripting.FileSystemObject")
Set objStreamout = objFSOut.CreateTextFile(fileout, 1, false)

objStreamout.WriteLine("Inicio: " & date() & " " & time())
do while objStream.AtEndOfStream <> true
        linea = objStream.readline
        BorrarUsuario(linea)
loop
objStreamout.write "Final: " & date() & " " & time()
objStreamout.close()
Set objStreamout = nothing
Set objFSOut = nothing

objStream.close()
Set objStream = nothing
Set objFSO = nothing

msgbox "Listo el pollo" & vbCrLf & "Pelada la gallina"

sub BorrarUsuario(cuenta)

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

textolog = "El usuario " & linea & " fue eliminado del AD." & vbCrLf
textolog = textolog & "========================================" & vbCrLf

else

textolog = "El usuario " & linea & " no fue encontrado." & vbCrLf
textolog = textolog & "========================================" & vbCrLf

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
                        vUsuario = Replace(rs(0),"GC://","")
                        FindUser=vUsuario
                 else
                        FindUser = "Not Found"
                end if
        end if
        cn.close
end function
