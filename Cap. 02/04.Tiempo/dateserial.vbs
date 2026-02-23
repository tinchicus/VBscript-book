dim f

fecha = date
f = split(fecha,"/")
dia = dateserial(f(2) + 5, f(1) - 2, f(0) + 13)

msgbox formatdatetime(dia, vbLongDate)
