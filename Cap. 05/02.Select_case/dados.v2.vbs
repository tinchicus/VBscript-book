randomize timer

dim dado
dim texto

msgbox "Tirar el dado"

dado = int(rnd * 6) + 1

select case dado
        case 1
                texto = "Salio el numero 1"
        case 2
                texto = "Salio el numero 2"
        case 3
                texto = "Salio el numero 3"
        case 4
                texto = "Salio el numero 4"
        case 5
                texto = "Salio el numero 5"
        case else
                texto = "Salto el numero 6"
end select

wscript.echo(texto)
