num = 13
base = 10

wscript.echo(int(-(LogBase(num, base))))
wscript.echo(fix(-(LogBase(num, base))))

Function LogBase (numero, base)
        LogBase = Log(numero) / Log(Base)
end Function
