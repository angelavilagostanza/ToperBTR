<%@ CODEPAGE=65001 LANGUAGE="VBScript" %>
<!--#include virtual="/include/bootstrap.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
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
<title>ToperBTR - Cambiar mi contraseña</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Cambiar mi contraseña</h2>
    <% If EsCadenaNoVacia(errorMsg) Then %>
    <p class="flash-error"><%= Server.HTMLEncode(errorMsg) %></p>
    <% End If %>
    <form method="post" action="/modules/usuarios/cambio_password_do.asp" class="formulario-vertical">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <label>Contraseña actual
            <input type="password" name="passwordActual" required>
        </label>
        <label>Nueva contraseña
            <input type="password" name="passwordNueva" minlength="10" required>
        </label>
        <label>Repita la nueva contraseña
            <input type="password" name="passwordNueva2" minlength="10" required>
        </label>
        <button type="submit">Cambiar contraseña</button>
    </form>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
