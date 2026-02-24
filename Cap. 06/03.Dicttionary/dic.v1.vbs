dim menu
dim miDic

texto = "Elige una accion: " & vbCrLf
texto = texto & "Si para Nuevo elemento" & vbCrLf
texto = texto & "No para Mostrar elementos" & vbCrLf
texto = texto & "Cancelar para borrar elementos"

Set miDic = CreateObject("Scripting.Dictionary")

sub nuevoElemento()

        dim valor
        dim clave
        clave = inputbox("Ingresa un nombre para la clave")
        if miDic.exists(clave) then
                msgbox "Ya existe esta clave"
                exit sub
        end if
        valor = inputbox("Ingresa el valor para la clave")
        miDic.add clave, valor

end sub

sub mostrarElementos()

        dim texto
        dim claves
        dim valores

        claves = miDic.keys
        valores = miDic.items
        for i = 0 to miDic.count-1
                texto = texto & claves(i) & ": " & valores(i) & vbCrLf

        next
        msgbox texto

end sub

do
        menu = msgbox(texto,3,"Diccionarios")   
        select case menu
                case 6
                        nuevoElemento()
                case 7
                        mostrarElementos()
                case else
                        miDic.removeall()
                        exit do
        end select
loop
