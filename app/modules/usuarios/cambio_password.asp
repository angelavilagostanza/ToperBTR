<!--#include virtual="/include/bootstrap.asp"-->
<%
Call RequiereSesion()

Dim token
token = GenerarTokenCsrf()

Dim errorMsg
errorMsg = Request.QueryString("error")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Cambiar mi contrasena</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Cambiar mi contrasena</h2>
    <% If EsCadenaNoVacia(errorMsg) Then %>
    <p class="flash-error"><%= Server.HTMLEncode(errorMsg) %></p>
    <% End If %>
    <form method="post" action="/modules/usuarios/cambio_password_do.asp" class="formulario-vertical">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <label>Contrasena actual
            <input type="password" name="passwordActual" required>
        </label>
        <label>Nueva contrasena
            <input type="password" name="passwordNueva" minlength="10" required>
        </label>
        <label>Repita la nueva contrasena
            <input type="password" name="passwordNueva2" minlength="10" required>
        </label>
        <button type="submit">Cambiar contrasena</button>
    </form>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
