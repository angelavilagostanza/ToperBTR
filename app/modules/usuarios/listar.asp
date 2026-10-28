<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/usuarios_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequierePerfil(Array(CodigoPerfil(conn, "Seguridad")))

Dim token
token = GenerarTokenCsrf()

Dim filtroLogin, filtroNombre, filtroApellidos, filtroPerfil, filtroEstado
filtroLogin = Trim(Request.QueryString("login"))
filtroNombre = Trim(Request.QueryString("nombre"))
filtroApellidos = Trim(Request.QueryString("apellidos"))
filtroPerfil = Request.QueryString("perfil")
filtroEstado = Request.QueryString("estado")

Dim limite, pagina, totalRegistros, totalPaginas, offsetReg

Select Case CStr(Request.QueryString("limite"))
    Case "100":  limite = 100
    Case "1000": limite = 1000
    Case Else:   limite = 25
End Select

pagina = 1
If IsNumeric(Request.QueryString("pagina")) Then
    If CLng(Request.QueryString("pagina")) > 1 Then
        pagina = CLng(Request.QueryString("pagina"))
    End If
End If

totalRegistros = ContarUsuarios(conn, filtroLogin, filtroNombre, filtroApellidos, filtroPerfil, filtroEstado)

totalPaginas = Int((totalRegistros + limite - 1) / limite)

If totalPaginas < 1 Then totalPaginas = 1

If pagina > totalPaginas Then pagina = totalPaginas

offsetReg = (pagina - 1) * limite

Dim rsUsuarios, rsPerfiles, rsEstados
Set rsUsuarios = BuscarUsuarios(conn, filtroLogin, filtroNombre, filtroApellidos, filtroPerfil, filtroEstado, limite, offsetReg)

Set rsPerfiles = ListarPerfiles(conn)
Set rsEstados = ListarEstadosUsuario(conn)

Function UrlPagina(nuevaPagina, nuevoLimite)

    Dim qs

    qs = "limite=" & nuevoLimite
    qs = qs & "&pagina=" & nuevaPagina

    If EsCadenaNoVacia(filtroLogin) Then
        qs = qs & "&login=" & Server.URLEncode(filtroLogin)
    End If

    If EsCadenaNoVacia(filtroNombre) Then
        qs = qs & "&nombre=" & Server.URLEncode(filtroNombre)
    End If

    If EsCadenaNoVacia(filtroApellidos) Then
        qs = qs & "&apellidos=" & Server.URLEncode(filtroApellidos)
    End If

    If EsCadenaNoVacia(filtroPerfil) Then
        qs = qs & "&perfil=" & Server.URLEncode(filtroPerfil)
    End If

    If EsCadenaNoVacia(filtroEstado) Then
        qs = qs & "&estado=" & Server.URLEncode(filtroEstado)
    End If

    UrlPagina = "/modules/usuarios/listar.asp?" & qs

End Function

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

    <nav class="modulo-nav">
        <span class="modulo-nav-item activo">Buscar</span>
        <a href="/modules/usuarios/alta.asp" class="modulo-nav-item">+ Nuevo usuario</a>
    </nav>

    <form method="get" action="/modules/usuarios/listar.asp" class="formulario-busqueda">
		<div class="formulario-fila">
        <label>Login
            <input type="text" name="login" maxlength="50" value="<%= Server.HTMLEncode(filtroLogin) %>">
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
		<label>Registros
			<select name="limite" onchange="this.form.submit();">
				<option value="25"<% If limite = 25 Then Response.Write " selected" %>>25</option>
				<option value="100"<% If limite = 100 Then Response.Write " selected" %>>100</option>
				<option value="1000"<% If limite = 1000 Then Response.Write " selected" %>>1000</option>
			</select>
		</label>
		<button type="submit">Filtrar</button>
		</div>
    </form>

    <form method="post" action="/modules/usuarios/confirmar_accion.asp">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <input type="hidden" name="volver" value="<%= Server.HTMLEncode(UrlPagina(pagina, limite)) %>">
		<div class="paginacion">
			<span class="paginacion-info">
				<%= totalRegistros %> resultado<% If totalRegistros <> 1 Then Response.Write "s" %>
				&mdash; P&aacute;gina <%= pagina %> de <%= totalPaginas %>
			</span>
			<nav class="paginacion-nav" aria-label="Paginaci&oacute;n">
				<% If pagina > 1 Then %>
				<a href="<%= UrlPagina(pagina - 1, limite) %>" class="paginacion-link">&larr; Anterior</a>
				<% Else %>
				<span class="paginacion-link deshabilitado">&larr; Anterior</span>
				<% End If %>
				<% If pagina < totalPaginas Then %>
				<a href="<%= UrlPagina(pagina + 1, limite) %>" class="paginacion-link">Siguiente &rarr;</a>
				<% Else %>
				<span class="paginacion-link deshabilitado">Siguiente &rarr;</span>
				<% End If %>
			</nav>
		</div>
        <table class="tabla-datos">
        <thead>
        <tr><th></th><th>Login</th><th>Nombre</th><th>Apellidos</th><th>Perfil</th><th>Estado</th><th>Fecha Alta</th><th>Ultima conexion</th><th></th><th></th></tr>
        </thead>
        <tbody>
        <% Do While Not rsUsuarios.EOF %>
        <tr>
            <td><input type="checkbox" name="seleccion" value="<%= rsUsuarios("COD_INDICE_USUARIO") %>"></td>
            <td><%= Server.HTMLEncode(rsUsuarios("COD_USUARIO") & "") %></td>
			<td><%= Server.HTMLEncode(rsUsuarios("NOMBRE") & "") %></td>
			<td><%= Server.HTMLEncode(rsUsuarios("APELLIDOS") & "") %></td>
			<td><%= Server.HTMLEncode(rsUsuarios("NOMBRE_PERFIL") & "") %></td>
			<td><%= Server.HTMLEncode(rsUsuarios("DESCRIPCION_ESTADO") & "") %></td>
			<td><%= SiVerdadero(IsDate(rsUsuarios("FECHA_ALTA")), rsUsuarios("FECHA_ALTA"), "&mdash;") %></td>
			<td><%= SiVerdadero(IsDate(rsUsuarios("FECHA_ULTIMA_CONEXION")), rsUsuarios("FECHA_ULTIMA_CONEXION"), "&mdash;") %></td>
            <td>
                <a href="/modules/usuarios/reset_password_admin.asp?cod=<%= rsUsuarios("COD_INDICE_USUARIO") %>" class="cabecera-icono" title="Cambiar contraseña">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                        <circle cx="8" cy="15" r="4"/><path d="M12 11l7-7m-3 0l3 3"/>
                    </svg>
                    <span class="cabecera-icono-label">Cambiar contrasena</span>
                </a>
            </td>
            <td>
                <a href="/modules/usuarios/editar.asp?cod=<%= rsUsuarios("COD_INDICE_USUARIO") %>" class="cabecera-icono" title="Modificar usuario">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                        <path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4L16.5 3.5z"/>
                    </svg>
                    <span class="cabecera-icono-label">Modificar</span>
                </a>
            </td>
        </tr>
        <% rsUsuarios.MoveNext
        Loop %>
        </tbody>
        </table>
		<%
		' Los tres botones se muestran siempre, independientemente del filtro de Estado
		' activo: la lista puede mezclar usuarios en distintos estados (sobre todo con
		' Estado=Todos), y aplicar una accion a un usuario al que no le corresponde es
		' un no-op seguro (ver guardas AND ESTADO <> ? en DarBajaUsuariosLote,
		' DesbloquearUsuariosLote y BloquearUsuariosLote, usuarios_dal.asp) — nunca
		' "resucita" a un usuario dado de baja ni rompe nada en un usuario ya en el
		' estado destino.
		%>
		<button type="submit" name="accion" value="baja">
			Dar de baja seleccionados
		</button>

		<button type="submit" name="accion" value="bloqueo">
			Bloquear seleccionados
		</button>

		<button type="submit" name="accion" value="desbloqueo">
			Desbloquear seleccionados
		</button>
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
