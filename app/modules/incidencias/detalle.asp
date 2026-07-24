<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/incidencias_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "LECTURA_INCIDENCIAS"))

If Not IsNumeric(Request.QueryString("cod")) Then
    conn.Close
    Response.Redirect "/errors/error.asp"
    Response.End
End If

Dim codIndIncidencia
codIndIncidencia = CLng(Request.QueryString("cod"))

Dim inc
Set inc = ObtenerIncidenciaPorIndice(conn, codIndIncidencia)
If inc Is Nothing Then
    conn.Close
    Response.Redirect "/errors/error.asp"
    Response.End
End If

Dim estadoVigente, puedeComentar, rsComentarios
estadoVigente = EstadoVigenteIncidencia(conn, codIndIncidencia)
puedeComentar = TienePermiso(Session("AccionesPermitidas"), CodigoAccion(conn, "ESCRITURA_INCIDENCIAS"))
Set rsComentarios = ObtenerComentarios(conn, codIndIncidencia)

Dim token, errorMsg
token = GenerarTokenCsrf()
errorMsg = Request.QueryString("error")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Incidencia <%= Server.HTMLEncode(inc("CodIncidencia")) %></title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Incidencia <%= Server.HTMLEncode(inc("CodIncidencia")) %></h2>

    <table class="tabla-datos">
        <tr><th>Tipo</th><td><%= Server.HTMLEncode(inc("TipoIncidencia")) %></td></tr>
        <tr><th>Categoria</th><td><%= Server.HTMLEncode(inc("Categoria")) %></td></tr>
        <tr><th>Operador</th><td><%= Server.HTMLEncode(inc("Operador")) %></td></tr>
        <tr><th>Estado</th><td><%= Server.HTMLEncode(estadoVigente) %></td></tr>
        <tr><th>Causa</th><td><%= Server.HTMLEncode(inc("Causa")) %></td></tr>
        <tr><th>Otros datos</th><td><%= Server.HTMLEncode(inc("OtrosDatos")) %></td></tr>
        <tr><th>Fecha creacion</th><td><%= inc("FechaCreacion") %></td></tr>
        <tr><th>Fecha resolucion</th><td><%= inc("FechaResolucion") & "" %></td></tr>
        <tr><th>Vinculada a</th><td>
            <% If inc("TipoObjeto") = "S" And EsCadenaNoVacia(inc("CodSolicitud")) Then %>
            Solicitud <a href="/modules/solicitudes/detalle.asp?cod=<%= inc("CodObjeto") %>"><%= Server.HTMLEncode(inc("CodSolicitud")) %></a>
            <% ElseIf inc("TipoObjeto") = "F" And EsCadenaNoVacia(inc("NombreFichero")) Then %>
            Fichero <%= Server.HTMLEncode(inc("NombreFichero")) %> (consulta detallada disponible en el Entregable 5)
            <% Else %>
            &mdash;
            <% End If %>
        </td></tr>
    </table>

    <h3>Comentarios</h3>
    <% If rsComentarios.EOF Then %>
    <p>Sin comentarios todavia.</p>
    <% Else %>
    <table class="tabla-datos">
    <thead><tr><th>Usuario</th><th>Fecha</th><th>Comentario</th></tr></thead>
    <tbody>
    <% Do While Not rsComentarios.EOF %>
    <tr>
        <td><%= Server.HTMLEncode(rsComentarios("COD_USUARIO")) %></td>
        <td><%= rsComentarios("FECHA_CREACION") %></td>
        <td><%= Server.HTMLEncode(rsComentarios("COMENTARIO")) %></td>
    </tr>
    <% rsComentarios.MoveNext
    Loop %>
    </tbody>
    </table>
    <% End If %>

    <% If puedeComentar Then %>
    <% If EsCadenaNoVacia(errorMsg) Then %>
    <p class="flash-error"><%= Server.HTMLEncode(errorMsg) %></p>
    <% End If %>
    <form method="post" action="/modules/incidencias/comentario_do.asp" class="formulario-vertical">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <input type="hidden" name="cod" value="<%= codIndIncidencia %>">
        <label>Nuevo comentario
            <textarea name="comentario" maxlength="500" rows="4" required></textarea>
        </label>
        <button type="submit">Anadir comentario</button>
    </form>
    <% End If %>

    <p><a href="/modules/incidencias/buscar.asp">Volver a la busqueda</a></p>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
<%
rsComentarios.Close
conn.Close
Set conn = Nothing
%>
