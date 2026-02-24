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
        else
                exit do
        end if
        a = a + 1               
loop

wscript.echo("Tus numeros fueron: " & texto & " - hasta el ciclo: " & a)
