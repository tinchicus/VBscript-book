dim entrada

entrada = inputbox("Ingresa una direccion de correo")

Set er = new RegExp
with er
        .Pattern = "^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"
        .IgnoreCase = False
        .Global = False
end with

if er.test(entrada) then
        msgbox entrada & " es un email valido"
else
        msgbox entrada & " no es un email valido"
end if

set er = nothing
