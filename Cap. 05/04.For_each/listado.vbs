dim lista
dim texto

lista = "Martin Miranda,Enzo Tortore,DarkZero Aleman,Marta Gargaglione,Ariel Polizzi"

listilla = split(lista,",")

for each nombre in listilla
	texto = texto & nombre & vbCrLf
next

msgbox texto,,"Ejemplo de for generico"
