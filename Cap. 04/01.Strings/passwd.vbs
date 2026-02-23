randomize timer

dim texto

for a = 1 to 8
        caracter = int(rnd * 25) + 97
        texto = texto + chr(caracter)
next

wscript.echo texto 
