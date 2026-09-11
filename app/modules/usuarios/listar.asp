<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/usuarios_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequierePerfil(Array(CodigoPerfil(conn, "Seguridad")))

Dim token
token = GenerarTokenCsrf()

Dim filtroLogin, filtroNombre, filtroApellidos, filtroPerfil, filtroEstado
filtroLogin = Trim(Request.Form("login"))
filtroNombre = Trim(Request.Form("nombre"))
filtroApellidos = Trim(Request.Form("apellidos"))
filtroPerfil = Request.Form("perfil")
filtroEstado = Request.Form("estado")

Dim rsUsuarios, rsPerfiles, rsEstados
Set rsUsuarios = BuscarUsuarios(conn, filtroLogin, filtroNombre, filtroApellidos, filtroPerfil, filtroEstado)
Set rsPerfiles = ListarPerfiles(conn)
Set rsEstados = ListarEstadosUsuario(conn)
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Usuarios</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Usuarios</h2>
    <p><a href="/modules/usuarios/alta.asp">Dar de alta un usuario nuevo</a></p>

    <form method="post" action="/modules/usuarios/listar.asp" class="formulario-busqueda">
        <label>Login
            <input type="text" name="login" maxlength="15" value="<%= Server.HTMLEncode(filtroLogin) %>">
        </label>
        <label>Nombre
            <input type="text" name="nombre" maxlength="50" value="<%= Server.HTMLEncode(filtroNombre) %>">
        </label>
        <label>Apellidos
            <input type="text" name="apellidos" maxlength="200" value="<%= Server.HTMLEncode(filtroApellidos) %>">
        </label>
        <label>Perfil
            <select name="perfil">
                <option value="-1">Todos</option>
                <% Do While Not rsPerfiles.EOF %>
                <option value="<%= rsPerfiles("COD_PERFIL") %>"<% If CStr(filtroPerfil) = CStr(rsPerfiles("COD_PERFIL")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsPerfiles("NOMBRE")) %></option>
                <% rsPerfiles.MoveNext
                Loop %>
            </select>
        </label>
        <label>Estado
            <select name="estado">
                <option value="-1">Todos</option>
                <% Do While Not rsEstados.EOF %>
                <option value="<%= rsEstados("ESTADO") %>"<% If CStr(filtroEstado) = CStr(rsEstados("ESTADO")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsEstados("DESCRIPCION")) %></option>
                <% rsEstados.MoveNext
                Loop %>
            </select>
        </label>
        <button type="submit">Filtrar</button>
    </form>

    <form method="post" action="/modules/usuarios/confirmar_accion.asp">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <table class="tabla-datos">
        <thead>
        <tr><th></th><th>Login</th><th>Nombre</th><th>Apellidos</th><th>Perfil</th><th>Estado</th><th></th></tr>
        </thead>
        <tbody>
        <% Do While Not rsUsuarios.EOF %>
        <tr>
            <td><input type="checkbox" name="seleccion" value="<%= rsUsuarios("COD_INDICE_USUARIO") %>"></td>
            <td><%= Server.HTMLEncode(rsUsuarios("COD_USUARIO")) %></td>
            <td><%= Server.HTMLEncode(rsUsuarios("NOMBRE")) %></td>
            <td><%= Server.HTMLEncode(rsUsuarios("APELLIDOS")) %></td>
            <td><%= Server.HTMLEncode(rsUsuarios("NOMBRE_PERFIL")) %></td>
            <td><%= Server.HTMLEncode(rsUsuarios("DESCRIPCION_ESTADO")) %></td>
            <td><a href="/modules/usuarios/reset_password_admin.asp?cod=<%= rsUsuarios("COD_INDICE_USUARIO") %>">Cambiar contrasena</a></td>
        </tr>
        <% rsUsuarios.MoveNext
        Loop %>
        </tbody>
        </table>
        <button type="submit" name="accion" value="baja">Dar de baja seleccionados</button>
        <button type="submit" name="accion" value="desbloqueo">Desbloquear seleccionados</button>
    </form>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
<%
rsUsuarios.Close
rsPerfiles.Close
rsEstados.Close
conn.Close
Set conn = Nothing
%>
