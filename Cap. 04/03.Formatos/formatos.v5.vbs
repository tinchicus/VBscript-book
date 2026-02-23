dim texto

moneda = -128568.922345

moneda = formatcurrency(moneda,2,true,true)
texto = "FormatCurrency: " & vbCrLf & moneda

numero = formatnumber(moneda, 3, true, false)
texto = texto & vbCrLf & "FormatNumber: " & vbCrLf & numero

valor = "0,124"
valor = formatpercent(valor)
texto = texto & vbCrLf & "FormatPercent: " & vbCrLf & valor

fecha = formatdatetime("06/12/19", vbLongDate)
hora = formatdatetime("12:15:40", 4)
texto = texto & vbCrLf & "FormatDateTime: " & vbCrLf & fecha
texto = texto & vbCrLf & hora

texto = texto & vbCrLf & "String: " & vbCrLf & string(10, 65)
texto = texto & vbCrLf & "Space: " & vbCrLf & chr(64) & space(10) & chr(64)
mensaje = "    Este es un texto de prueba     "
texto = texto & vbCrLf & "Trim: " & vbCrLf & trim(mensaje)

msgbox texto,,"Distintos formatos"
