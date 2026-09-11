<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/usuarios_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequierePerfil(Array(CodigoPerfil(conn, "Seguridad")))

Dim token
token = GenerarTokenCsrf()

Dim rsPerfiles
Set rsPerfiles = ListarPerfiles(conn)

Dim errorMsg
errorMsg = Request.QueryString("error")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Alta de usuario</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Alta de usuario</h2>
    <% If EsCadenaNoVacia(errorMsg) Then %>
    <p class="flash-error"><%= Server.HTMLEncode(errorMsg) %></p>
    <% End If %>
    <form method="post" action="/modules/usuarios/alta_do.asp" class="formulario-vertical">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <label>Login
            <input type="text" name="usuario" maxlength="15" required>
        </label>
        <label>Nombre
            <input type="text" name="nombre" maxlength="50" required>
        </label>
        <label>Apellidos
            <input type="text" name="apellidos" maxlength="200" required>
        </label>
        <label>Contrasena
            <input type="password" name="password" minlength="10" maxlength="100" required>
        </label>
        <label>Repita la contrasena
            <input type="password" name="password2" minlength="10" maxlength="100" required>
        </label>
        <label>Perfil
            <select name="perfil" required>
                <option value="">-- Seleccione --</option>
                <% Do While Not rsPerfiles.EOF %>
                <option value="<%= rsPerfiles("COD_PERFIL") %>"><%= Server.HTMLEncode(rsPerfiles("NOMBRE")) %></option>
                <% rsPerfiles.MoveNext
                Loop %>
            </select>
        </label>
        <button type="submit">Crear usuario</button>
    </form>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
<%
rsPerfiles.Close
conn.Close
Set conn = Nothing
%>
