texto = "Texto de prueba para usar Join"

a = split(texto, chr(32))

texto = join(a, "_")

wscript.echo texto
