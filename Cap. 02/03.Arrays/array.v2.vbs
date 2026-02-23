dim arreglo(2,6)
dim texto

arreglo(0,0) = "Martin" 
arreglo(1,0) = "Miranda"
arreglo(0,1) = "Marta" 
arreglo(1,1) = "Gargaglione"
arreglo(0,2) = "Enzo" 
arreglo(1,2) = "Tortore"
arreglo(0,3) = "Javier" 
arreglo(1,3) = "Marcuzzi"
arreglo(0,4) = "Ariel" 
arreglo(1,4) = "Polizzi"
arreglo(0,5) = "Raul" 
arreglo(1,5) = "Picos"

for a = 0 to 5
        apellido = arreglo(1,a)
        nombre = arreglo(0,a)
        texto = texto & apellido &  ", " & nombre & "  " & chr(10)
next

msgbox texto,,"Devolucion del Array"
