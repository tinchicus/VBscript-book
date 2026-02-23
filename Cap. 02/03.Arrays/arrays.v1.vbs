dim arreglo(6)
dim texto

arreglo(0) = "Martin Miranda"
arreglo(1) = "Marta Gargaglione"
arreglo(2) = "Enzo Tortore"
arreglo(3) = "Javier Marcuzzi"
arreglo(4) = "Ariel Polizzi"
arreglo(5) = "Raul Picos"

for a = 0 to 5
        texto = texto & arreglo(a) & chr(10)
        msgbox texto
next
