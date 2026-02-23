lista = "tinchicus@gmail.com, a.b.com, marta@yahoo.com, mrbogusa@gmail.com, c.gmail.com, mirandma.ar.ibm.com"

l = split(lista, ", ")

valido = filter(l, "@")
novalido = filter(l, "@", false)

for a = 0 to ubound(valido)
        vale = vale & valido(a) & ", "
next

for a = 0 to ubound(valido)
        novale = novale & novalido(a) & ", "
next

texto = "Validos: " & vale & vbCrLf
texto = texto & "No validos: " & novale

wscript.echo texto
