set WshShell=WScript.CreateObject("WScript.Shell")
strAnswer = InputBox("Por favor ingrese un nombre para su archivo:")
WshShell.run "chrome.exe"
WScript.sleep 100
WshShell.sendkeys "tinchicus.com" 
WshShell.sendkeys "{ENTER}"
