dim correo
dim caracter
dim texto

correo = inputbox("Ingresa una direccion de correo")

for a = 1 to len(correo)
        caracter = mid(correo, a, 1)
        if caracter = chr(64) then
                texto = correo & " es valido"
                exit for
        end if
        i = i + 1
next

if i = len(correo) then texto = correo & " no es valido"

wscript.echo texto
