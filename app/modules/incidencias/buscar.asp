<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/incidencias_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "LECTURA_INCIDENCIAS"))

Dim fCod, fTipo, fEstado, fOperador, fCodSolicitud, fNombreFichero, fDesde, fHasta
fCod = Trim(Request.Form("codIncidencia"))
fTipo = Request.Form("tipo")
fEstado = Request.Form("estado")
fOperador = Request.Form("operador")
fCodSolicitud = Trim(Request.Form("codSolicitud"))
fNombreFichero = Trim(Request.Form("nombreFichero"))
fDesde = Trim(Request.Form("fechaDesde"))
fHasta = Trim(Request.Form("fechaHasta"))

Dim rsIncidencias, rsTipos, rsEstados, rsOperadores
Set rsIncidencias = BuscarIncidencias(conn, fCod, fTipo, fEstado, fOperador, fCodSolicitud, fNombreFichero, fDesde, fHasta)
Set rsTipos = ListarTiposIncidencia(conn)
Set rsEstados = ListarEstadosIncidencia(conn)
Set rsOperadores = ListarOperadores(conn)
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Incidencias</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Incidencias</h2>

    <form method="post" action="/modules/incidencias/buscar.asp" class="formulario-busqueda">
        <label>Codigo <input type="text" name="codIncidencia" value="<%= Server.HTMLEncode(fCod) %>"></label>
        <label>Tipo
            <select name="tipo">
                <option value="-1">Todos</option>
                <% Do While Not rsTipos.EOF %>
                <option value="<%= rsTipos("COD_TIPO_INCIDENCIA") %>"<% If CStr(fTipo) = CStr(rsTipos("COD_TIPO_INCIDENCIA")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsTipos("DESCRIPCION")) %></option>
                <% rsTipos.MoveNext
                Loop %>
            </select>
        </label>
        <label>Estado
            <select name="estado">
                <option value="-1">Todos</option>
                <% Do While Not rsEstados.EOF %>
                <option value="<%= rsEstados("COD_ESTADO") %>"<% If CStr(fEstado) = CStr(rsEstados("COD_ESTADO")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsEstados("DESCRIPCION")) %></option>
                <% rsEstados.MoveNext
                Loop %>
            </select>
        </label>
        <label>Operador
            <select name="operador">
                <option value="-1">Todos</option>
                <% Do While Not rsOperadores.EOF %>
                <option value="<%= rsOperadores("COD_OPERADOR") %>"<% If CStr(fOperador) = CStr(rsOperadores("COD_OPERADOR")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsOperadores("NOMBRE")) %></option>
                <% rsOperadores.MoveNext
                Loop %>
            </select>
        </label>
        <label>Codigo de solicitud <input type="text" name="codSolicitud" value="<%= Server.HTMLEncode(fCodSolicitud) %>"></label>
        <label>Nombre de fichero <input type="text" name="nombreFichero" value="<%= Server.HTMLEncode(fNombreFichero) %>"></label>
        <label>Creacion desde <input type="date" name="fechaDesde" value="<%= Server.HTMLEncode(fDesde) %>"></label>
        <label>Creacion hasta <input type="date" name="fechaHasta" value="<%= Server.HTMLEncode(fHasta) %>"></label>
        <button type="submit">Filtrar</button>
    </form>

    <table class="tabla-datos">
    <thead>
    <tr><th>Codigo</th><th>Tipo</th><th>Operador</th><th>Estado</th><th>Vinculada a</th><th>Creacion</th></tr>
    </thead>
    <tbody>
    <% Do While Not rsIncidencias.EOF %>
    <tr>
        <td><a href="/modules/incidencias/detalle.asp?cod=<%= rsIncidencias("COD_IND_INCIDENCIA") %>"><%= Server.HTMLEncode(rsIncidencias("COD_INCIDENCIA")) %></a></td>
        <td><%= Server.HTMLEncode(rsIncidencias("TIPO_INCIDENCIA")) %></td>
        <td><%= Server.HTMLEncode(rsIncidencias("OPERADOR")) %></td>
        <td><%= Server.HTMLEncode(rsIncidencias("ESTADO")) %></td>
        <td>
            <% If rsIncidencias("TIPO_OBJETO") = "S" Then %>
            Solicitud <%= Server.HTMLEncode(rsIncidencias("COD_SOLICITUD") & "") %>
            <% ElseIf rsIncidencias("TIPO_OBJETO") = "F" Then %>
            Fichero <%= Server.HTMLEncode(rsIncidencias("NOMBRE_FICHERO") & "") %>
            <% Else %>
            &mdash;
            <% End If %>
        </td>
        <td><%= rsIncidencias("FECHA_CREACION") %></td>
    </tr>
    <% rsIncidencias.MoveNext
    Loop %>
    </tbody>
    </table>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
<%
rsIncidencias.Close
rsTipos.Close
rsEstados.Close
rsOperadores.Close
conn.Close
Set conn = Nothing
%>
