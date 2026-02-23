dim total

sub p_Sumar(a, b)
        total = a + b
end sub

function f_Sumar(a, b)
        f_Sumar = a + b
end function

p_Sumar 3, 4
wscript.echo total

total = f_Sumar(4, 5)
wscript.echo total
