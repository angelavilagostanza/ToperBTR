<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/solicitudes_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "LECTURA_SOLICITUDES"))

If Not IsNumeric(Request.QueryString("cod")) Then
    conn.Close
    Response.Redirect "/errors/error.asp"
    Response.End
End If

Dim codIndiceSol
codIndiceSol = CLng(Request.QueryString("cod"))

Dim sol
Set sol = ObtenerSolicitudPorIndice(conn, codIndiceSol)
If sol Is Nothing Then
    conn.Close
    Response.Redirect "/errors/error.asp"
    Response.End
End If

Dim puedeEscribir, esCancelable
puedeEscribir = TienePermiso(Session("AccionesPermitidas"), CodigoAccion(conn, "ESCRITURA_SOLICITUDES"))
esCancelable = EsSolicitudCancelable(conn, codIndiceSol)

Dim rsHistorico, rsConfirmaciones, numIncidencias
Set rsHistorico = ObtenerHistoricoEstados(conn, codIndiceSol)
Set rsConfirmaciones = ObtenerConfirmaciones(conn, codIndiceSol)
numIncidencias = ContarIncidenciasAsociadas(conn, codIndiceSol)

Dim token
token = GenerarTokenCsrf()
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Solicitud <%= Server.HTMLEncode(sol("CodSolicitud")) %></title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Solicitud <%= Server.HTMLEncode(sol("CodSolicitud")) %></h2>

    <table class="tabla-datos">
        <tr><th>Tipo</th><td><%= Server.HTMLEncode(sol("Tipo")) %></td></tr>
        <tr><th>IMEI</th><td><%= Server.HTMLEncode(sol("Imei")) %></td></tr>
        <tr><th>MSISDN</th><td><%= Server.HTMLEncode(sol("Msisdn")) %></td></tr>
        <tr><th>Numero de denuncia</th><td><%= Server.HTMLEncode(sol("NumDenuncia")) %></td></tr>
        <tr><th>Fecha de robo</th><td><%= sol("FechaRobo") & "" %></td></tr>
        <tr><th>Fecha de denuncia</th><td><%= sol("FechaDenuncia") & "" %></td></tr>
        <tr><th>Fecha de creacion</th><td><%= sol("FechaCreacion") %></td></tr>
        <tr><th>Creada por</th><td><%= Server.HTMLEncode(sol("UsuarioCreador")) %></td></tr>
    </table>

    <h3>Cliente</h3>
    <table class="tabla-datos">
        <tr><th>NIF/CIF</th><td><%= Server.HTMLEncode(sol("NumIdentificacion")) %></td></tr>
        <tr><th>Nombre</th><td><%= Server.HTMLEncode(sol("NombreCliente")) %> <%= Server.HTMLEncode(sol("PrimerApellido")) %> <%= Server.HTMLEncode(sol("SegundoApellido")) %></td></tr>
    </table>

    <h3>Historico de estados</h3>
    <table class="tabla-datos">
    <thead><tr><th>Estado</th><th>Desde</th><th>Hasta</th></tr></thead>
    <tbody>
    <% Do While Not rsHistorico.EOF %>
    <tr>
        <td><%= Server.HTMLEncode(rsHistorico("DESCRIPCION")) %></td>
        <td><%= rsHistorico("FECHA_INICIO") %></td>
        <td><%= rsHistorico("FECHA_FIN") & "" %></td>
    </tr>
    <% rsHistorico.MoveNext
    Loop %>
    </tbody>
    </table>

    <h3>Confirmaciones</h3>
    <% If rsConfirmaciones.EOF Then %>
    <p>Sin confirmaciones todavia.</p>
    <% Else %>
    <table class="tabla-datos">
    <thead><tr><th>Operador</th><th>Tipo</th><th>Fecha</th></tr></thead>
    <tbody>
    <% Do While Not rsConfirmaciones.EOF %>
    <tr>
        <td><%= Server.HTMLEncode(rsConfirmaciones("OPERADOR")) %></td>
        <td><%= Server.HTMLEncode(rsConfirmaciones("TIPO_CONFIRMACION")) %></td>
        <td><%= rsConfirmaciones("FECHA_CONFIRMACION") %></td>
    </tr>
    <% rsConfirmaciones.MoveNext
    Loop %>
    </tbody>
    </table>
    <% End If %>

    <% If numIncidencias > 0 Then %>
    <p class="flash-info">Esta solicitud tiene <%= numIncidencias %> incidencia(s) asociada(s) (la consulta detallada estara disponible en el Entregable 4).</p>
    <% End If %>

    <% If puedeEscribir And esCancelable Then %>
    <form method="post" action="/modules/solicitudes/cancelar_confirmar.asp">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <input type="hidden" name="seleccion" value="<%= codIndiceSol %>">
        <button type="submit">Cancelar esta solicitud</button>
    </form>
    <% End If %>

    <p><a href="/modules/solicitudes/consulta.asp">Volver a la busqueda</a></p>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
<%
rsHistorico.Close
rsConfirmaciones.Close
conn.Close
Set conn = Nothing
%>
