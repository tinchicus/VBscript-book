dim texto

moneda = -128568.922345

moneda = formatcurrency(moneda,2,true,true)
texto = "FormatCurrency: " & vbCrLf & moneda
numero = formatnumber(moneda, 3, true, false)
texto = texto & vbCrLf & "FormatNumber: " & vbCrLf & numero

msgbox texto,,"Distintos formatos"
