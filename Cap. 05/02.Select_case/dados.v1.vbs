randomize timer

dim dado
dim texto

msgbox "Tirar el dado"

dado = int(rnd * 6) + 1

if dado = 1 then texto = "Salio el numero 1"
if dado = 2 then texto = "Salio el numero 2"
if dado = 3 then texto = "Salio el numero 3"
if dado = 4 then texto = "Salio el numero 4"
if dado = 5 then texto = "Salio el numero 5"
if dado = 6 then texto = "Salio el numero 6"

wscript.echo(texto)
