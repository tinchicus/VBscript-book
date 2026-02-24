randomize timer

dim dado
dim texto

for a = 1 to 6

        msgbox "Tirar el dado"
        dado = int(rnd * 6) + 1
        texto = texto & " " & dado

next

wscript.echo("Tus numeros fueron: " & texto)
