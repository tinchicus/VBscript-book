dim texto

moneda = -128568.922345

moneda = formatcurrency(moneda,2,true,true)
texto = "FormatCurrency: " & vbCrLf & moneda
numero = formatnumber(moneda, 3, true, false)
texto = texto & vbCrLf & "FormatNumber: " & vbCrLf & numero
valor = "0,124"
valor = formatpercent(valor)
texto = texto & vbCrLf & "FormatPercent: " & vbCrLf & valor

msgbox texto,,"Distintos formatos"
