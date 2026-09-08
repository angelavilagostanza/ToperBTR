<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/solicitudes_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "LECTURA_SOLICITUDES"))

Dim puedeEscribir
puedeEscribir = TienePermiso(Session("AccionesPermitidas"), CodigoAccion(conn, "ESCRITURA_SOLICITUDES"))

Dim token
token = GenerarTokenCsrf()

Dim fCod, fImei, fMsisdn, fTipo, fEstado, fIdent, fNombre, fDesde, fHasta
fCod = Trim(Request.Form("codSolicitud"))
fImei = Trim(Request.Form("imei"))
fMsisdn = Trim(Request.Form("msisdn"))
fTipo = Request.Form("tipo")
fEstado = Request.Form("estado")
fIdent = Trim(Request.Form("identCliente"))
fNombre = Trim(Request.Form("nombreCliente"))
fDesde = Trim(Request.Form("fechaDesde"))
fHasta = Trim(Request.Form("fechaHasta"))

Dim rsSolicitudes, rsTipos, rsEstados
'Set rsSolicitudes = BuscarSolicitudes(conn, fCod, fImei, fMsisdn, fTipo, fEstado, fIdent, fNombre, fDesde, fHasta)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    Set rsSolicitudes = BuscarSolicitudes( _
        conn, _
        fCod, _
        fImei, _
        fMsisdn, _
        fTipo, _
        fEstado, _
        fIdent, _
        fNombre, _
        fDesde, _
        fHasta)

Else

    Set rsSolicitudes = EjecutarConsulta( _
        conn, _
        "SELECT S.COD_INDICE_SOL, S.COD_SOLICITUD, T.DESCRIPCION AS TIPO, " & _
        "S.IMEI, S.MSISDN, S.FECHA_CREACION, " & _
        "C.NOMBRE AS NOMBRE_CLIENTE, C.PRIMER_APELLIDO, '' AS ESTADO_VIGENTE " & _
        "FROM SOLICITUD S " & _
        "INNER JOIN TIPO_SOLICITUD_REF T ON S.COD_TIPO_SOLICITUD = T.COD_TIPO_SOLICITUD " & _
        "INNER JOIN CLIENTE C ON S.COD_CLIENTE = C.COD_CLIENTE " & _
        "WHERE 1 = 0", _
        Array())

End If
Set rsTipos = ListarTiposSolicitud(conn)
Set rsEstados = ListarEstadosSolicitud(conn)
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Solicitudes</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Solicitudes</h2>
    <% If puedeEscribir Then %>
    <p><a href="/modules/solicitudes/pre_alta.asp">Nueva solicitud</a></p>
    <% End If %>

    <form method="post" action="/modules/solicitudes/consulta.asp" class="formulario-busqueda">
        <label>Codigo <input type="text" name="codSolicitud" value="<%= Server.HTMLEncode(fCod) %>"></label>
        <label>IMEI <input type="text" name="imei" value="<%= Server.HTMLEncode(fImei) %>"></label>
        <label>MSISDN <input type="text" name="msisdn" value="<%= Server.HTMLEncode(fMsisdn) %>"></label>
        <label>Tipo
            <select name="tipo">
                <option value="-1">Todos</option>
                <% Do While Not rsTipos.EOF %>
                <option value="<%= rsTipos("COD_TIPO_SOLICITUD") %>"<% If CStr(fTipo) = CStr(rsTipos("COD_TIPO_SOLICITUD")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsTipos("DESCRIPCION")) %></option>
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
        <label>NIF/CIF cliente <input type="text" name="identCliente" value="<%= Server.HTMLEncode(fIdent) %>"></label>
        <label>Nombre cliente <input type="text" name="nombreCliente" value="<%= Server.HTMLEncode(fNombre) %>"></label>
        <label>Creacion desde <input type="date" name="fechaDesde" value="<%= Server.HTMLEncode(fDesde) %>"></label>
        <label>Creacion hasta <input type="date" name="fechaHasta" value="<%= Server.HTMLEncode(fHasta) %>"></label>
        <button type="submit">Filtrar</button>
    </form>

    <form method="post" action="/modules/solicitudes/cancelar_confirmar.asp">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <table class="tabla-datos">
        <thead>
        <tr>
            <% If puedeEscribir Then %><th></th><% End If %>
            <th>Codigo</th><th>Tipo</th><th>IMEI</th><th>MSISDN</th><th>Cliente</th><th>Estado</th><th>Creacion</th>
        </tr>
        </thead>
        <tbody>
        <% Do While Not rsSolicitudes.EOF %>
        <tr>
            <% If puedeEscribir Then %>
            <td>
                <% If EsSolicitudCancelable(conn, rsSolicitudes("COD_INDICE_SOL")) Then %>
                <input type="checkbox" name="seleccion" value="<%= rsSolicitudes("COD_INDICE_SOL") %>">
                <% End If %>
            </td>
            <% End If %>
            <td><a href="/modules/solicitudes/detalle.asp?cod=<%= rsSolicitudes("COD_INDICE_SOL") %>"><%= Server.HTMLEncode(rsSolicitudes("COD_SOLICITUD")) %></a></td>
            <td><%= Server.HTMLEncode(rsSolicitudes("TIPO")) %></td>
            <td><%= Server.HTMLEncode(rsSolicitudes("IMEI")) %></td>
            <td><%= Server.HTMLEncode(rsSolicitudes("MSISDN") & "") %></td>
            <td><%= Server.HTMLEncode(rsSolicitudes("CLIENTE") & "") %></td>
			<td><%= Server.HTMLEncode(rsSolicitudes("ESTADO_VIGENTE")) %></td>
            <td><%= rsSolicitudes("FECHA_CREACION") %></td>
        </tr>
        <% rsSolicitudes.MoveNext
        Loop %>
        </tbody>
        </table>
        <% If puedeEscribir Then %>
        <button type="submit">Cancelar solicitudes seleccionadas</button>
        <% End If %>
    </form>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
<%
rsSolicitudes.Close
rsTipos.Close
rsEstados.Close
conn.Close
Set conn = Nothing
%>
