dim miDic
dim texto

set miDic = CreateObject("Scripting.Dictionary")

miDic.add "tinchicus","Martin Miranda"
miDic.add "etortore","Enzo Tortore"
miDic.add "daleman","DarkZero Aleman"
miDic.add "gargaglm","Marta Gargaglione"
miDic.add "marcuzzj","Javier Marcuzzi"
miDic.add "polizzia","Ariel Polizzi"

for each usuario in miDic.keys
	texto = texto & usuario & ": " & miDic(usuario) & vbCrLf
next

msgbox texto,,"Ejemplo de Diccionario"
