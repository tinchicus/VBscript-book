dim arreglo(1,5)
dim texto
dim lista
dim total
lista = "Martin Miranda Marta Gargaglione Enzo Tortore Javier Marcuzzi Ariel Polizzi Raul Picos"
l = split(lista, chr(32))
b = 0
c = 0
for a = 0 to ubound(l)
        if b > 1 then 
                b = 0
                c = c + 1
        end if
        arreglo(b,c) = l(a)
        b = b + 1
next
for a = 0 to 5
        apellido = arreglo(1,a)
        nombre = arreglo(0,a)
        texto = texto & apellido &  ", " & nombre & "  " & chr(10)
next

msgbox texto,,"Devolucion del Array"
