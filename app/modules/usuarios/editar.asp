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

Dim rsPerfiles
Set rsPerfiles = ListarPerfiles(conn)

Dim errorMsg
errorMsg = Request.QueryString("error")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Modificar usuario</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Modificar usuario <code><%= Server.HTMLEncode(u("CodUsuario")) %></code></h2>
    <% If EsCadenaNoVacia(errorMsg) Then %>
    <p class="flash-error"><%= Server.HTMLEncode(errorMsg) %></p>
    <% End If %>
    <form method="post" action="/modules/usuarios/editar_do.asp" class="formulario-vertical">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <input type="hidden" name="cod" value="<%= u("CodIndiceUsuario") %>">
        <label>Nombre
            <input type="text" name="nombre" maxlength="50" value="<%= Server.HTMLEncode(u("Nombre")) %>" required>
        </label>
        <label>Apellidos
            <input type="text" name="apellidos" maxlength="200" value="<%= Server.HTMLEncode(u("Apellidos")) %>" required>
        </label>
        <label>Perfil
            <select name="perfil" required>
                <option value="">-- Seleccione --</option>
                <% Do While Not rsPerfiles.EOF %>
                <option value="<%= rsPerfiles("COD_PERFIL") %>"<% If CLng(rsPerfiles("COD_PERFIL")) = u("CodPerfil") Then Response.Write " selected" %>><%= Server.HTMLEncode(rsPerfiles("NOMBRE")) %></option>
                <% rsPerfiles.MoveNext
                Loop %>
            </select>
        </label>
        <div class="acciones-form">
            <button type="submit">Guardar cambios</button>
            <a href="/modules/usuarios/listar.asp" class="btn btn-secundario">Cancelar</a>
        </div>
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
