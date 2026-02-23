private const miConst = 100

wscript.echo("El valor de miConst: " & miConst)
otraConst
miConst = cambiarConst()

sub otraConst()
        const miConst = 50
        wscript.echo("El valor de otraConst: " & miConst)
end sub

function cambiarConst()
        cambiarConst = 10
end function
