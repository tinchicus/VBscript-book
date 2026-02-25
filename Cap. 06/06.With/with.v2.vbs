class compu
	public tipo
	public cpu
	private so
	public property let SistOper(s)
		so = s
	end property
	public function getOs()
		getOs = so
	end function
end class

dim texto
set pc = new compu
with pc
	.tipo = "Notebook"
	.cpu = "Intel"
	.SistOper = "Windows 10"
end with
texto = "Tipo: " & pc.tipo & vbCrLf
texto = texto & "CPU: " & pc.cpu & vbCrLf
texto = texto & "SO: " & pc.getOs() & vbCrLf
msgbox texto,,"Ejemplo con with"
