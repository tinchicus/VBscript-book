dim texto
dim a
dim i

for a = Asc("A") to Asc("Z")
        texto = texto & " " & chr(a) & "=" & a
        i = i + 1
        if i = 6 then
                texto = texto & vbCrLf
                i = 0
        end if
next

wscript.echo texto 
