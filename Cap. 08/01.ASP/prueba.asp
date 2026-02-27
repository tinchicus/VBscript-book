<html>
<head><title><% response.write("Titulo en ASP")%></title>
</head>
<body>
<% response.write("Esto fue escrito en ASP") %>
<% for a = 1 to 10
        response.write("<p>El valor de a es: " & a & "</p>")
   next %>
<% response.write("El bucle anterior tambien") %>
</body>
</html>
