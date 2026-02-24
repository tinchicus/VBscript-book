txt="Este es un hermoso dia"
Set objReg=CreateObject("vbscript.regexp")
objReg.Pattern="e"
wscript.echo objReg.Replace(txt,"##")
