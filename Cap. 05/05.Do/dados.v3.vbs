randomize timer

dim dado
dim texto
dim a

a = 0

do
        consulta = msgbox("Tirar el dado", vbOkCancel)
        if consulta = 1 then
                dado = int(rnd * 6) + 1
                texto = texto & " " & dado
                msgbox "Tu numero es " & dado
                a = a + 1
        else
                exit do
        end if
loop until a >= 6

wscript.echo("Tus numeros fueron: " & texto & " - hasta el ciclo: " & a)
