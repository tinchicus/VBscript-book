dim columnas
dim filas
dim texto

columnas = inputbox("Ingresa el numero de columnas")
filas = inputbox("Ingresa el numero de filas")

for a = 1 to filas
        for b =  1 to columnas
                c = a * b
                texto = texto & " " & c         
        next
        texto = texto & chr(10)
next

wscript.echo(texto)
