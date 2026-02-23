dim texto

moneda = -128568.922345

moneda = formatcurrency(moneda,2,true,true)
texto = "FormatCurrency: " & vbCrLf & moneda

msgbox texto,,"Distintos formatos"
