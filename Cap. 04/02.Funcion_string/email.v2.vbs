dim correo
dim caracter
dim texto

correo = inputbox("Ingresa una direccion de correo")

pos = instr(correo, chr(64))

if pos = 0 then
        texto = correo & " no es valido"
else
        texto = correo & " es valido"
end if

wscript.echo texto
