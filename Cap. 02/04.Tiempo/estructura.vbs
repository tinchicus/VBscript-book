dim fecha

fecha = date

dia = day(fecha)
mes = month(fecha)
ano = year(fecha)
diaSem = weekday(fecha, vbUseSystem)
nombreMes = monthname(mes)
diaNom = weekdayname(diaSem)

texto = "Day: " & dia & vbCrLf
texto = texto & "Month: " & mes & vbCrLf
texto = texto & "Year: " & ano & vbCrLf
texto = texto & "WeekDay: " & diaSem & vbCrLf
texto = texto & "MonthName: " & nombreMes & vbCrLf
texto = texto & "WeekDayName: " & diaNom & vbCrLf

msgbox texto,,"Todos los tipos de estructuras"
