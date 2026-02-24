ON ERROR RESUME NEXT

for a=1 to 13
        err.raise a
        msgbox "Error #" & Cstr(err.number) & " / " & err.description
        err.clear
next
