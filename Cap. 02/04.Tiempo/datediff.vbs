dim nacim
dim texto

nacim = inputbox("ingresa tu cumpleaños:")
texto = "Lleva: " & vbCrLf
texto = texto & datediff("yyyy", nacim, now) & " años" & vbCrLf
texto = texto & datediff("q", nacim, now) & " trimestres" & vbCrLf
texto = texto & datediff("d", nacim, now) & " dias" & vbCrLf
texto = texto & datediff("h", nacim, now) & " horas" & vbCrLf
texto = texto & datediff("n", nacim, now) & " minutos" & vbCrLf
texto = texto & datediff("s", nacim, now) & " segundos" & vbCrLf

msgbox texto,,"Ud. lleva este tiempo sobre la tierra"
