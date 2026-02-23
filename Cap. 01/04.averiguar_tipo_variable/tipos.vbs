dim a

texto = texto & typename(a) & ": " & vartype(a) & chr(10)
a = null
texto = texto & typename(a) & ": " & vartype(a) & chr(10)
a = "hola, mundo!"
texto = texto & typename(a) & ": " & vartype(a) & chr(10)
a = 100
texto = texto & typename(a) & ": " & vartype(a) & chr(10)

mens=MsgBox(texto,64,"Ejemplo de distintos tipos")
