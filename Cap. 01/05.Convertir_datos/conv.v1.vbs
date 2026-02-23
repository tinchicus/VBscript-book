dim var1
dim var2
dim total

var1 = "11"
var2 = "22"

total = var1 + var2
wscript.echo total & " - " & typename(total)
total = Cint(var1) + Cint(var2)
wscript.echo total & " - " & typename(total)
