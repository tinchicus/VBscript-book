Const ADS_GROUP_TYPE_LOCAL_GROUP = &h4
Const ADS_GROUP_TYPE_GLOBAL_GROUP = &h2
Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &h8
Const ADS_GROUP_TYPE_SECURITY_ENABLED = &h80000000
Const titulo = "Alta de grupo v. 1.0"

nombre = inputbox("Ingresa el nombre",titulo)
texto = "Elige el rango del grupo:" & vbCrLf
texto = texto & "1)Dominio Local" & vbCrLf
texto = texto & "2)Global" & vbCrLf
texto = texto & "3)Universal"
rango = inputbox(texto,titulo,1)
texto = "Elige el tipo del grupo:" & vbCrLf
texto = texto & "1)Seguridad" & vbCrLf
texto = texto & "2)Distribucion"
tipo = inputbox(texto, titulo,1)

Set objOU = GetObject("LDAP://cn=users,dc=laboratorio,dc=local")
Set objGroup = objOU.Create("Group", "cn=" & nombre)
objGroup.Put "sAMAccountName", replace(nombre,chr(32),"_",1,100)

select case rango

        case 3
                if tipo = 2 then
                        objGroup.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP
                else
                        objGroup.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP or _
                        ADS_GROUP_TYPE_SECURITY_ENABLED
                end if 
                
        case 2
                if tipo = 2 then
                        objGroup.Put "groupType", ADS_GROUP_TYPE_GLOBAL_GROUP
                else
                        objGroup.Put "groupType", ADS_GROUP_TYPE_GLOBAL_GROUP or _
                        ADS_GROUP_TYPE_SECURITY_ENABLED
                end if
        case else
                if tipo = 2 then
                        objGroup.Put "groupType", ADS_GROUP_TYPE_LOCAL_GROUP            
                else
                        objGroup.Put "groupType", ADS_GROUP_TYPE_LOCAL_GROUP or _
                        ADS_GROUP_TYPE_SECURITY_ENABLED
                end if 
                

end select

objGroup.SetInfo

msgbox "Se acabo el script"
