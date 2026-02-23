num = 13
base = 10

texto = "Num. original: " & LogBase(num, base) & chr(10)
texto = texto & "Num. con 5 digitos: " & round(LogBase(num, base),5) & chr(10)
texto = texto & "Num. con 0 digitos: " & round(LogBase(num, base))

wscript.echo(texto)

Function LogBase (numero, base)
        LogBase = Log(numero) / Log(Base)
end Function
