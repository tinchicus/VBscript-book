Const cTitulo = "Convertidor de path a OU v.1.0"
dim path, OU, final

path = inputbox("Ingresa el Path",cTitulo)

OU = ConvertirOU(path)

final = inputbox("Tu OU convertido",cTitulo,OU)

function ConvertirOU(ByVal camino)
        dim c, texto
	  c = split(camino,"/")
        for a=ubound(c) to 1 step -1
                texto = texto & "ou=" & c(a) & ","
        next
        texto = mid(texto,1,len(texto)-1)
        ConvertirOU = texto
end function
