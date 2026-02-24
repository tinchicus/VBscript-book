dim miDic
dim texto
dim user

set miDic = CreateObject("Scripting.Dictionary")

miDic.add "tinchicus","Martin Miranda"
miDic.add "etortore","Enzo Tortore"
miDic.add "daleman","DarkZero Aleman"
miDic.add "gargaglm","Marta Gargaglione"
miDic.add "marcuzzj","Javier Marcuzzi"
miDic.add "polizzia","Ariel Polizzi" 

user = inputbox("Ingresa un usuario a buscar:","Ejemplo de Diccionario")

for each usuario in miDic.keys
	if user = usuario then
		texto = usuario & ": " & miDic(usuario)
	end if
next
if texto = "" then texto = user & " no fue encontrado"

msgbox texto,,"Ejemplo de Diccionario"
