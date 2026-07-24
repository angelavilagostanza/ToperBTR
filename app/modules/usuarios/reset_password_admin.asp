<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/usuarios_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequierePerfil(Array(CodigoPerfil(conn, "Seguridad")))

If Not IsNumeric(Request.QueryString("cod")) Then
    conn.Close
    Response.Redirect "/errors/error.asp"
    Response.End
End If

Dim u
Set u = ObtenerUsuarioPorIndice(conn, CLng(Request.QueryString("cod")))

If u Is Nothing Then
    conn.Close
    Response.Redirect "/errors/error.asp"
    Response.End
End If

Dim token
token = GenerarTokenCsrf()

Dim errorMsg
errorMsg = Request.QueryString("error")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Cambiar contrasena de usuario</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Cambiar contrasena de <%= Server.HTMLEncode(u("CodUsuario")) %></h2>
    <% If EsCadenaNoVacia(errorMsg) Then %>
    <p class="flash-error"><%= Server.HTMLEncode(errorMsg) %></p>
    <% End If %>
    <form method="post" action="/modules/usuarios/reset_password_admin_do.asp" class="formulario-vertical">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <input type="hidden" name="cod" value="<%= u("CodIndiceUsuario") %>">
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
<%
conn.Close
Set conn = Nothing
%>
