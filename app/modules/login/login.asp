<!--#include virtual="/include/bootstrap.asp"-->
<%
Dim token, mensajeError
token = GenerarTokenCsrf()
mensajeError = Request.QueryString("error")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Acceso</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body class="pagina-login">
<div class="caja-login">
    <img src="/images/ToperBTR_Logo_peque_v3.png" alt="ToperBTR" class="login-logo">
    <p class="login-subtitulo">Bloqueo de Terminales Robados</p>
    <% If mensajeError <> "" Then %>
    <p class="mensaje-error"><%= Server.HTMLEncode(mensajeError) %></p>
    <% End If %>
    <form method="post" action="/modules/login/login_do.asp" autocomplete="off">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <label for="usuario">Usuario</label>
        <input type="text" id="usuario" name="usuario" maxlength="15" required autofocus>
        <label for="password">Contraseña</label>
        <input type="password" id="password" name="password" maxlength="100" required>
        <button type="submit">Entrar</button>
    </form>
</div>
</body>
</html>
