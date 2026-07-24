<!--#include virtual="/include/bootstrap.asp"-->
<%
Call RequiereSesion()
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Inicio</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Bienvenido, <%= Server.HTMLEncode(Session("Nombre")) %></h2>
    <p>Seleccione una opcion en el menu.</p>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
